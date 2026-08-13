"""
在后台线程中执行批量简历转换，并向GUI发送结构化事件。

本模块不创建Tk窗口，因此可以独立进行离线单元测试。
"""

from __future__ import annotations

import io
import threading
from collections.abc import Callable
from contextlib import redirect_stdout
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from main import BatchResult, ReadabilityChecker, TemplateOption
from resume_pipeline import ResumeConverter


EventEmitter = Callable[["WorkerEvent"], None]
AccessResolver = Callable[[Path, str], str]
Converter = Callable[[str, str, str, bool], str | None]


@dataclass(frozen=True)
class WorkerEvent:
    """后台任务发送给GUI的一条状态事件。"""

    kind: str
    payload: dict[str, Any]


class _EventWriter(io.TextIOBase):
    """把业务层print内容按行转发到GUI日志。"""

    def __init__(self, emit: EventEmitter) -> None:
        self._emit = emit
        self._buffer = ""

    def write(self, value: str) -> int:
        self._buffer += value
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line.strip():
                self._emit(WorkerEvent("log", {"message": line.rstrip()}))
        return len(value)

    def flush(self) -> None:
        if self._buffer.strip():
            self._emit(
                WorkerEvent("log", {"message": self._buffer.rstrip()})
            )
        self._buffer = ""


class GuiBatchWorker:
    """按文件顺序执行转换，并只在文件边界响应安全停止。"""

    def __init__(
        self,
        files: list[Path],
        output_directory: Path,
        template: TemplateOption,
        emit: EventEmitter,
        resolve_access: AccessResolver,
        stop_event: threading.Event,
        converter: Converter = ResumeConverter,
        readability_checker: ReadabilityChecker | None = None,
    ) -> None:
        self.files = files
        self.output_directory = output_directory
        self.template = template
        self.emit = emit
        self.resolve_access = resolve_access
        self.stop_event = stop_event
        self.converter = converter
        self.readability_checker = (
            readability_checker or _check_file_readability
        )

    def run(self) -> BatchResult:
        """执行任务并返回汇总；未处理文件统一记为跳过。"""

        succeeded: list[Path] = []
        failed: list[Path] = []
        skipped: list[Path] = []
        total = len(self.files)
        stopped_early = False

        self.emit(WorkerEvent("started", {"total": total}))

        for index, source_path in enumerate(self.files, start=1):
            if self.stop_event.is_set():
                skipped.extend(self.files[index - 1 :])
                stopped_early = True
                break

            self.emit(
                WorkerEvent(
                    "file_started",
                    {
                        "index": index,
                        "total": total,
                        "path": source_path,
                    },
                )
            )

            access_action = self._resolve_file_access(source_path)
            if access_action == "skip":
                skipped.append(source_path)
                self._emit_file_finished(index, total, source_path, "skipped")
                continue
            if access_action == "stop":
                skipped.extend(self.files[index - 1 :])
                stopped_early = True
                break

            writer = _EventWriter(self.emit)
            try:
                with redirect_stdout(writer):
                    result = self.converter(
                        str(source_path),
                        str(self.output_directory),
                        self.template.template_tag,
                        self.template.with_photo,
                    )
            except Exception as error:
                self.emit(
                    WorkerEvent(
                        "log",
                        {"message": f"[未处理异常] {error}"},
                    )
                )
                result = None
            finally:
                writer.flush()

            if result:
                succeeded.append(Path(result).resolve())
                status = "succeeded"
            else:
                failed.append(source_path)
                status = "failed"
            self._emit_file_finished(index, total, source_path, status)

            if self.stop_event.is_set() and index < total:
                skipped.extend(self.files[index:])
                stopped_early = True
                break

        result = BatchResult(
            total=total,
            succeeded=tuple(succeeded),
            failed=tuple(failed),
            skipped=tuple(skipped),
            stopped_early=stopped_early,
        )
        self.emit(WorkerEvent("completed", {"result": result}))
        return result

    def _resolve_file_access(self, source_path: Path) -> str:
        while True:
            readable, message = self.readability_checker(source_path)
            if readable:
                return "ready"
            action = self.resolve_access(source_path, message)
            if action in {"retry", "skip", "stop"}:
                if action == "retry":
                    continue
                return action

    def _emit_file_finished(
        self,
        index: int,
        total: int,
        source_path: Path,
        status: str,
    ) -> None:
        self.emit(
            WorkerEvent(
                "file_finished",
                {
                    "index": index,
                    "total": total,
                    "path": source_path,
                    "status": status,
                },
            )
        )


def _check_file_readability(path: Path) -> tuple[bool, str]:
    """实际读取一个字节，检测文件是否被其他程序独占。"""

    try:
        with path.open("rb") as file:
            file.read(1)
        return True, ""
    except (PermissionError, OSError) as error:
        return False, str(error)
