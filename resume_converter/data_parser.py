"""
原始简历解析流程入口。

处理流程：
1. 调用百度文档解析生成中间Markdown；
2. 调用内网大模型恢复Markdown结构；
3. 返回最终Markdown文件路径；
4. 过程数据统一写入process_data目录；
5. 每次处理结果追加到processing_records.xlsx；
6. 真正异常写入error_YYYYMMDD.txt；
7. 异常不再向终端抛出完整Traceback。

本模块暴露DataParser函数。
"""

from __future__ import annotations

import traceback
from pathlib import Path

from utils.baidu_parser import BaiduParser
from utils.error_logging import write_error_log
from utils.processing_records import (
    TRACKING_WORKBOOK_NAME,
    append_document_record,
)
from utils.restore_markdown import (
    SUPPORTED_SUFFIXES,
    RestoreMarkdown,
)


PROCESS_DATA_DIRNAME = "process_data"
DOCUMENT_PARSING_DIRNAME = "document_parsing"


def DataParser(
    input_file_path: str,
    output_dir: str,
) -> str | None:
    """
    将原始简历解析为Markdown并恢复结构。

    路径规则：

        原始文件：
        任意目录/XXX.pdf、XXX.docx、XXX.jpg等

        百度中间文件：
        output_dir/process_data/document_parsing/XXX.md

        最终文件：
        output_dir/process_data/document_parsing/XXX_restored.md

    Args:
        input_file_path:
            原始PDF、Word或图片文件路径。

        output_dir:
            中间文件和最终文件的输出目录。

    Returns:
        成功时返回最终Markdown文件绝对路径。

        失败时返回None，并将详细异常写入：
        output_dir/process_data/error_YYYYMMDD.txt
    """

    source_path = Path(
        input_file_path
    )

    output_root = Path(
        output_dir
    )

    output_directory = (
        output_root
        / PROCESS_DATA_DIRNAME
    )

    document_parsing_directory = (
        output_directory
        / DOCUMENT_PARSING_DIRNAME
    )

    tracking_workbook = (
        output_directory
        / TRACKING_WORKBOOK_NAME
    )

    input_file = (
        document_parsing_directory
        / f"{source_path.stem}.md"
    )

    output_file = (
        document_parsing_directory
        / f"{source_path.stem}_restored.md"
    )

    try:
        if not source_path.exists():
            raise FileNotFoundError(
                "找不到输入文件："
                f"{source_path.resolve()}"
            )

        if source_path.suffix.lower() not in SUPPORTED_SUFFIXES:
            if source_path.suffix.lower() == ".doc":
                raise ValueError(
                    "不支持旧版DOC文件，请先转换为DOCX格式。"
                )

            raise ValueError(
                "不支持的输入文件格式："
                f"{source_path.suffix.lower() or '无后缀'}；"
                "支持PDF、DOCX、JPG、JPEG和PNG。"
            )

        document_parsing_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        generated_input_file = BaiduParser(
            input_file_path=str(source_path),
            output_dir=str(document_parsing_directory),
        )

        if (
            Path(generated_input_file).resolve()
            != input_file.resolve()
        ):
            raise RuntimeError(
                "百度解析输出路径与预期路径不一致："
                f"实际={generated_input_file}，"
                f"预期={input_file.resolve()}"
            )

        result_file = RestoreMarkdown(
            input_file_path=str(source_path),
            input_file=str(input_file),
            output_file=str(output_file),
        )

        append_document_record(
            workbook_path=tracking_workbook,
            source_path=source_path,
            baidu_markdown_path=input_file,
            restored_markdown_path=Path(result_file),
            status="成功",
        )

        print(f"  [文档解析完成] {result_file}")

        return result_file

    except Exception as error:
        error_file = write_error_log(
            output_dir=output_directory,
            input_file_path=source_path,
            stage="文档解析",
        )

        try:
            append_document_record(
                workbook_path=tracking_workbook,
                source_path=source_path,
                baidu_markdown_path=input_file,
                restored_markdown_path=output_file,
                status="失败",
                error_message=str(error),
            )
        except Exception:
            # 跟踪表写入失败不能掩盖原始错误，补充记录到文本日志。
            with error_file.open(
                "a",
                encoding="utf-8",
            ) as file:
                file.write(
                    "处理记录写入失败：\n"
                    f"{traceback.format_exc()}\n"
                )

        print(f"  [文档解析失败] {error}")

        print(
            "  详细错误已写入："
            f"{error_file.resolve()}"
        )

        return None
