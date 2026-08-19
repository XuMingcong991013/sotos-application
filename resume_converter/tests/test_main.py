"""
交互式批量处理入口的离线测试。
"""

from __future__ import annotations

import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from main import (
    TEMPLATE_OPTIONS,
    _find_duplicate_stems,
    discover_resume_files,
    process_folder,
)


class MainBatchTests(unittest.TestCase):
    """验证文件发现、模板配置和批处理失败隔离。"""

    def test_discovers_supported_files_only_in_current_directory(self) -> None:
        """只扫描当前层级，并忽略DOC、文本及子目录文件。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            directory = Path(temp)
            for name in (
                "a.pdf",
                "b.DOCX",
                "c.jpg",
                "d.jpeg",
                "e.png",
                "~$opened.docx",
                "old.doc",
                "note.txt",
            ):
                (directory / name).write_bytes(b"test")
            nested = directory / "nested"
            nested.mkdir()
            (nested / "inside.pdf").write_bytes(b"test")

            supported, ignored = discover_resume_files(directory)

            self.assertEqual(
                [path.name for path in supported],
                ["a.pdf", "b.DOCX", "c.jpg", "d.jpeg", "e.png"],
            )
            self.assertEqual(
                [path.name for path in ignored],
                ["note.txt", "old.doc", "~$opened.docx"],
            )

    def test_template_options_map_photo_modes_without_branching(self) -> None:
        """数字菜单通过配置映射到标签和照片模式。"""

        self.assertEqual(TEMPLATE_OPTIONS["1"].template_tag, "SOTOS")
        self.assertTrue(TEMPLATE_OPTIONS["1"].with_photo)
        self.assertEqual(TEMPLATE_OPTIONS["2"].template_tag, "SOTOS")
        self.assertFalse(TEMPLATE_OPTIONS["2"].with_photo)
        self.assertEqual(TEMPLATE_OPTIONS["3"].template_tag, "优族")
        self.assertFalse(TEMPLATE_OPTIONS["3"].with_photo)
        self.assertEqual(TEMPLATE_OPTIONS["4"].template_tag, "奇瑞")
        self.assertFalse(TEMPLATE_OPTIONS["4"].with_photo)

    def test_process_folder_continues_after_single_failure(self) -> None:
        """一份失败不影响下一份，并正确统计结果。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            directory = Path(temp)
            files = [directory / "a.pdf", directory / "b.pdf"]
            for path in files:
                path.write_bytes(b"test")
            calls: list[tuple[str, str, str, bool]] = []

            def fake_converter(
                input_file: str,
                output_dir: str,
                template_tag: str,
                with_photo: bool,
            ) -> str | None:
                calls.append(
                    (input_file, output_dir, template_tag, with_photo)
                )
                if Path(input_file).name == "a.pdf":
                    return None
                return str(directory / "b_标准简历.docx")

            with redirect_stdout(io.StringIO()) as output:
                result = process_folder(
                    files,
                    directory / "output",
                    TEMPLATE_OPTIONS["2"],
                    converter=fake_converter,
                )

            self.assertEqual(result.total, 2)
            self.assertEqual(len(result.succeeded), 1)
            self.assertEqual(result.failed, (files[0],))
            self.assertEqual(result.skipped, ())
            self.assertEqual(len(calls), 2)
            self.assertTrue(all(call[2] == "SOTOS" for call in calls))
            self.assertTrue(all(call[3] is False for call in calls))
            self.assertIn("正在处理 [1/2]：a.pdf", output.getvalue())
            self.assertIn("正在处理 [2/2]：b.pdf", output.getvalue())

    def test_duplicate_stems_are_detected_case_insensitively(self) -> None:
        """不同后缀但同名主体的文件会被识别为覆盖风险。"""

        files = [Path("候选人.pdf"), Path("候选人.DOCX"), Path("其他.png")]
        duplicates = _find_duplicate_stems(files)

        self.assertEqual(list(duplicates), ["候选人"])
        self.assertEqual(len(duplicates["候选人"]), 2)

    def test_locked_file_can_retry_and_then_process(self) -> None:
        """被占用文件关闭后选择重试，可以继续正常处理。"""

        source = Path("locked.docx")
        checks = iter(
            [(False, "Permission denied"), (True, "")]
        )
        converter_calls: list[str] = []

        def fake_converter(*args) -> str:
            converter_calls.append(args[0])
            return "locked_标准简历.docx"

        with redirect_stdout(io.StringIO()):
            result = process_folder(
                [source],
                Path("output"),
                TEMPLATE_OPTIONS["1"],
                converter=fake_converter,
                input_reader=lambda _: "r",
                readability_checker=lambda _: next(checks),
            )

        self.assertEqual(len(converter_calls), 1)
        self.assertEqual(len(result.succeeded), 1)
        self.assertEqual(result.skipped, ())

    def test_locked_file_can_be_skipped_without_calling_converter(self) -> None:
        """用户跳过被占用文件时不会调用任何外部服务。"""

        source = Path("locked.pdf")
        converter_calls: list[str] = []

        def fake_converter(*args) -> str:
            converter_calls.append(args[0])
            return "unexpected.docx"

        with redirect_stdout(io.StringIO()):
            result = process_folder(
                [source],
                Path("output"),
                TEMPLATE_OPTIONS["2"],
                converter=fake_converter,
                input_reader=lambda _: "s",
                readability_checker=lambda _: (
                    False,
                    "Permission denied",
                ),
            )

        self.assertEqual(converter_calls, [])
        self.assertEqual(result.skipped, (source,))
        self.assertEqual(result.failed, ())
