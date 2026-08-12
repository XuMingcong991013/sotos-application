"""
单份简历统一转换入口的离线测试。
"""

from __future__ import annotations

import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

import resume_pipeline
from resume_pipeline import ResumeConverter


class ResumePipelineTests(unittest.TestCase):
    """验证公共解析前置、模板分发和失败短路。"""

    def test_sotos_pipeline_returns_final_word(self) -> None:
        """SOTOS依次调用三个阶段并返回最终Word路径。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            temp_path = Path(temp)
            source_path = temp_path / "resume.pdf"
            output_dir = temp_path / "output"
            restored_path = output_dir / "resume_restored.md"
            json_path = output_dir / "resume_extracted.json"
            word_path = output_dir / "final_resumes" / "resume_标准简历.docx"
            source_path.write_bytes(b"pdf")

            with (
                patch.object(
                    resume_pipeline,
                    "DataParser",
                    return_value=str(restored_path),
                ) as data_parser,
                patch.object(
                    resume_pipeline,
                    "InformationExtractor",
                    return_value=str(json_path),
                ) as extractor,
                patch.object(
                    resume_pipeline,
                    "ResumeGenerator",
                    return_value=str(word_path),
                ) as generator,
            ):
                result = ResumeConverter(
                    str(source_path),
                    str(output_dir),
                    " sotos ",
                    True,
                )

            self.assertEqual(result, str(word_path))
            data_parser.assert_called_once_with(
                input_file_path=str(source_path),
                output_dir=str(output_dir),
            )
            extractor.assert_called_once_with(
                input_file_path=str(source_path),
                input_file=str(restored_path),
                output_dir=str(output_dir),
            )
            generator.assert_called_once_with(
                input_file=str(json_path),
                output_dir=str(output_dir),
                with_photo=True,
            )

    def test_document_parse_failure_stops_later_stages(self) -> None:
        """公共文档解析失败后不再调用模板专用阶段。"""

        with (
            patch.object(
                resume_pipeline,
                "DataParser",
                return_value=None,
            ),
            patch.object(
                resume_pipeline,
                "InformationExtractor",
            ) as extractor,
            patch.object(
                resume_pipeline,
                "ResumeGenerator",
            ) as generator,
        ):
            result = ResumeConverter(
                "missing.pdf",
                "output",
                "SOTOS",
                False,
            )

        self.assertIsNone(result)
        extractor.assert_not_called()
        generator.assert_not_called()

    def test_extraction_failure_stops_word_generation(self) -> None:
        """模板专用信息提取失败后不生成Word。"""

        with (
            patch.object(
                resume_pipeline,
                "DataParser",
                return_value="resume_restored.md",
            ),
            patch.object(
                resume_pipeline,
                "InformationExtractor",
                return_value=None,
            ),
            patch.object(
                resume_pipeline,
                "ResumeGenerator",
            ) as generator,
        ):
            result = ResumeConverter(
                "resume.pdf",
                "output",
                "SOTOS",
                False,
            )

        self.assertIsNone(result)
        generator.assert_not_called()

    def test_unknown_tag_is_logged_before_document_parsing(self) -> None:
        """未知标签不会消耗解析服务，并写入统一错误日志。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            output_dir = Path(temp) / "output"

            with (
                patch.object(resume_pipeline, "DataParser") as data_parser,
                redirect_stdout(io.StringIO()),
            ):
                result = ResumeConverter(
                    "resume.pdf",
                    str(output_dir),
                    "UNKNOWN",
                    False,
                )

            self.assertIsNone(result)
            data_parser.assert_not_called()
            error_file = next(
                (output_dir / "process_data").glob("error_*.txt")
            )
            error_text = error_file.read_text(encoding="utf-8")
            self.assertIn("处理阶段：标准简历统一流程", error_text)
            self.assertIn("不支持的模板标签：UNKNOWN", error_text)

    def test_invalid_photo_mode_is_rejected_before_document_parsing(self) -> None:
        """证件照模式不是布尔值时不调用文档解析服务。"""

        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            output_dir = Path(temp) / "output"

            with (
                patch.object(resume_pipeline, "DataParser") as data_parser,
                redirect_stdout(io.StringIO()),
            ):
                result = ResumeConverter(
                    "resume.pdf",
                    str(output_dir),
                    "SOTOS",
                    "False",  # type: ignore[arg-type]
                )

            self.assertIsNone(result)
            data_parser.assert_not_called()
            error_file = next(
                (output_dir / "process_data").glob("error_*.txt")
            )
            error_text = error_file.read_text(encoding="utf-8")
            self.assertIn(
                "with_photo必须是布尔值True或False",
                error_text,
            )
