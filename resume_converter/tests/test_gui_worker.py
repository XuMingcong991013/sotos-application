"""桌面GUI后台批处理和运行路径的离线测试。"""

from __future__ import annotations

import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

from gui import _selected_files_summary, _validate_selected_files
from gui_support import GuiBatchWorker, WorkerEvent
from main import TemplateOption
from utils.runtime_paths import (
    application_directory,
    application_environment_path,
    resource_path,
)


class GuiBatchWorkerTests(unittest.TestCase):
    """验证GUI线程外的批处理行为，不创建真实窗口。"""

    def setUp(self) -> None:
        self.template = TemplateOption("测试模板", "SOTOS", False)

    def test_selected_file_summary_supports_one_or_many_files(self) -> None:
        """输入框能清晰显示单文件路径和多文件数量。"""

        single = (Path(r"D:\resumes\one.pdf"),)
        many = tuple(Path(f"resume-{index}.pdf") for index in range(5))

        self.assertEqual(_selected_files_summary(single), str(single[0]))
        self.assertIn("已选择 5 份", _selected_files_summary(many))
        self.assertIn("……", _selected_files_summary(many))

    def test_selected_files_are_validated_without_folder_scanning(self) -> None:
        """GUI只处理用户选中的文件，并阻止同stem覆盖。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            directory = Path(temp)
            selected = directory / "chosen.pdf"
            unselected = directory / "not-chosen.pdf"
            selected.write_bytes(b"pdf")
            unselected.write_bytes(b"pdf")

            self.assertEqual(
                _validate_selected_files((selected,)),
                [selected.resolve()],
            )

            duplicate = directory / "chosen.docx"
            duplicate.write_bytes(b"docx")
            with self.assertRaisesRegex(ValueError, "文件名主体相同"):
                _validate_selected_files((selected, duplicate))

    def test_selected_files_reject_unsupported_format(self) -> None:
        """手工选到不支持格式时在调用外部API前拒绝。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            path = Path(temp) / "resume.doc"
            path.write_bytes(b"doc")
            with self.assertRaisesRegex(ValueError, "不支持的文件格式"):
                _validate_selected_files((path,))

    def test_reports_progress_and_continues_after_failure(self) -> None:
        """一份失败不会阻塞后续文件，并发送完整进度事件。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            directory = Path(temp)
            files = [directory / "a.pdf", directory / "b.docx"]
            for path in files:
                path.write_bytes(b"resume")

            events: list[WorkerEvent] = []

            def converter(input_path, output_dir, tag, with_photo):
                print("[阶段 1/3] 离线测试")
                if Path(input_path).name == "a.pdf":
                    return None
                result = Path(output_dir) / "result.docx"
                result.write_bytes(b"docx")
                return str(result)

            worker = GuiBatchWorker(
                files=files,
                output_directory=directory,
                template=self.template,
                emit=events.append,
                resolve_access=lambda path, message: "skip",
                stop_event=threading.Event(),
                converter=converter,
            )
            result = worker.run()

            self.assertEqual(len(result.failed), 1)
            self.assertEqual(len(result.succeeded), 1)
            self.assertFalse(result.stopped_early)
            self.assertEqual(
                [event.kind for event in events].count("file_finished"),
                2,
            )
            self.assertTrue(
                any(
                    event.kind == "log" and "阶段 1/3" in event.payload["message"]
                    for event in events
                )
            )

    def test_stop_request_is_honored_after_current_file(self) -> None:
        """停止标志在当前文件完成后生效，剩余文件记为跳过。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            directory = Path(temp)
            files = [directory / f"{index}.pdf" for index in range(3)]
            for path in files:
                path.write_bytes(b"resume")
            stop_event = threading.Event()
            calls: list[str] = []

            def converter(input_path, output_dir, tag, with_photo):
                calls.append(Path(input_path).name)
                stop_event.set()
                return str(Path(output_dir) / "result.docx")

            result = GuiBatchWorker(
                files=files,
                output_directory=directory,
                template=self.template,
                emit=lambda event: None,
                resolve_access=lambda path, message: "skip",
                stop_event=stop_event,
                converter=converter,
            ).run()

            self.assertEqual(calls, ["0.pdf"])
            self.assertEqual(len(result.succeeded), 1)
            self.assertEqual(len(result.skipped), 2)
            self.assertTrue(result.stopped_early)

    def test_locked_file_can_be_skipped(self) -> None:
        """文件被独占时使用GUI返回的选择，不调用转换器。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            path = Path(temp) / "locked.pdf"
            path.write_bytes(b"resume")
            calls: list[Path] = []

            result = GuiBatchWorker(
                files=[path],
                output_directory=Path(temp),
                template=self.template,
                emit=lambda event: None,
                resolve_access=lambda source, message: "skip",
                stop_event=threading.Event(),
                converter=lambda *args: calls.append(path),
                readability_checker=lambda source: (False, "占用中"),
            ).run()

            self.assertFalse(calls)
            self.assertEqual(result.skipped, (path,))


class RuntimePathTests(unittest.TestCase):
    """验证源码和PyInstaller模式下的路径不依赖当前工作目录。"""

    def test_source_resource_path_points_to_project(self) -> None:
        template = resource_path("templates", "standard_resume.docx")
        self.assertTrue(template.exists())
        self.assertEqual(application_directory(), Path.cwd().resolve())

    def test_frozen_paths_separate_external_config_and_resources(self) -> None:
        with patch("utils.runtime_paths.sys.frozen", True, create=True), patch(
            "utils.runtime_paths.sys.executable",
            r"C:\Program Files\ResumeTool\ResumeTool.exe",
        ), patch(
            "utils.runtime_paths.sys._MEIPASS",
            r"C:\Program Files\ResumeTool\_internal",
            create=True,
        ):
            self.assertEqual(
                application_directory(),
                Path(r"C:\Program Files\ResumeTool"),
            )
            self.assertEqual(
                resource_path("templates", "standard_resume.docx"),
                Path(
                    r"C:\Program Files\ResumeTool\_internal\templates"
                    r"\standard_resume.docx"
                ),
            )
            self.assertEqual(
                application_environment_path(),
                Path(r"C:\Program Files\ResumeTool\_internal\.env"),
            )


if __name__ == "__main__":
    unittest.main()
