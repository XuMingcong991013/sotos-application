"""
从人工样例中提炼不含候选人信息的标准简历模板。

该工具只在模板设计发生变化时使用，不是日常简历处理入口。
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn

from .generator import (
    SECTION_LABELS,
    _force_microsoft_yahei,
    _replace_xml_text,
    _section_elements,
    _set_document_defaults,
)


def build_standard_template(
    reference_file: str,
    output_file: str,
) -> str:
    """保留页眉页脚、品牌图片和六个分区标题，清除样例正文。"""

    reference_path = Path(reference_file)
    output_path = Path(output_file)

    if not reference_path.exists():
        raise FileNotFoundError(
            f"找不到模板参考文档：{reference_path.resolve()}"
        )

    document = Document(reference_path)
    sections = _section_elements(document)
    body = document._element.body

    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    for label in SECTION_LABELS:
        element: Any = deepcopy(sections[label])

        if label == "项目经验":
            _replace_xml_text(
                element,
                "Working Experience",
                "Project Experience",
            )

        body.insert(body.index(body.sectPr), element)

    _set_document_defaults(document)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    _force_microsoft_yahei(output_path)
    return str(output_path.resolve())
