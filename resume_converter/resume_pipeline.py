"""
单份原始简历到标准Word简历的统一流程入口。

文档解析流程对所有模板标签共用；信息提取和Word生成由模板标签分发。
当前仅支持SOTOS标签，后续新增模板时只需注册新的后半程处理函数。
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from data_parser import DataParser
from information_extraction import InformationExtractor
from resume_generation import ResumeGenerator
from utils.error_logging import write_error_log


PROCESS_DATA_DIRNAME = "process_data"
SOTOS_TEMPLATE_TAG = "SOTOS"

TemplateHandler = Callable[[str, str, str, bool], str | None]


def ResumeConverter(
    input_file_path: str,
    output_dir: str,
    template_tag: str,
    with_photo: bool,
) -> str | None:
    """
    将单份原始简历转换为指定标签的标准Word简历。

    Args:
        input_file_path: 原始PDF、DOCX、JPG、JPEG或PNG文件路径。
        output_dir: 过程数据和最终结果的输出根目录。
        template_tag: 模板标签；当前支持SOTOS，不区分大小写。
        with_photo: SOTOS是否生成证件照占位区域。

    Returns:
        成功时返回最终Word文件的绝对路径；任一阶段失败时返回None。
        真正异常追加写入output_dir/process_data/error_YYYYMMDD.txt。
    """

    source_path = Path(input_file_path)
    output_root = Path(output_dir)

    try:
        normalized_tag = _normalize_template_tag(template_tag)
        _validate_with_photo(with_photo)
        handler = TEMPLATE_HANDLERS.get(normalized_tag)

        if handler is None:
            supported = "、".join(sorted(TEMPLATE_HANDLERS))
            raise ValueError(
                f"不支持的模板标签：{normalized_tag}；"
                f"当前支持：{supported}。"
            )

        # 所有模板共享同一套百度解析和Markdown结构恢复流程。
        print("\n  [阶段 1/3] 文档解析与 Markdown 结构恢复")
        restored_file = DataParser(
            input_file_path=str(source_path),
            output_dir=str(output_root),
        )

        if not restored_file:
            return None

        print("\n  [阶段 2/3] 简历信息提取与证据校验")
        return handler(
            str(source_path),
            restored_file,
            str(output_root),
            with_photo,
        )

    except Exception as error:
        error_file = write_error_log(
            output_dir=output_root / PROCESS_DATA_DIRNAME,
            input_file_path=source_path,
            stage="标准简历统一流程",
        )
        print(f"标准简历流程失败：{error}")
        print(f"详细错误已写入：{error_file.resolve()}")
        return None


def _run_sotos_pipeline(
    input_file_path: str,
    restored_file: str,
    output_dir: str,
    with_photo: bool,
) -> str | None:
    """执行SOTOS标签专用的信息提取和Word生成。"""

    extracted_json = InformationExtractor(
        input_file_path=input_file_path,
        input_file=restored_file,
        output_dir=output_dir,
    )

    if not extracted_json:
        return None

    print("\n  [阶段 3/3] 标准 Word 与补充清单生成")
    return ResumeGenerator(
        input_file=extracted_json,
        output_dir=output_dir,
        with_photo=with_photo,
    )


def _normalize_template_tag(template_tag: str) -> str:
    """规范化模板标签，并拒绝空值或非字符串输入。"""

    if not isinstance(template_tag, str):
        raise TypeError("template_tag必须是字符串。")

    normalized = template_tag.strip().upper()

    if not normalized:
        raise ValueError("template_tag不能为空。")

    return normalized


def _validate_with_photo(with_photo: bool) -> None:
    """严格校验证件照模式，避免字符串等值被误判为True。"""

    if not isinstance(with_photo, bool):
        raise TypeError("with_photo必须是布尔值True或False。")


# 标签只控制信息提取和Word生成，DataParser始终位于分发之前。
TEMPLATE_HANDLERS: dict[str, TemplateHandler] = {
    SOTOS_TEMPLATE_TAG: _run_sotos_pipeline,
}
