"""
调用百度智能云办公文档识别接口，将简历转换为中间 Markdown。

PDF 和图片只使用 OCR 共享资源包支持的“办公文档识别”；DOCX 在本地
读取段落和表格，不调用百度服务。模块继续公开 BaiduParser，供现有流程复用。
"""

from __future__ import annotations

import base64
import os
import re
from pathlib import Path
from typing import Any
from urllib.parse import urlencode

import pymupdf
import requests
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from dotenv import load_dotenv

from utils.runtime_paths import application_environment_path


TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"
OFFICE_OCR_URL = (
    "https://aip.baidubce.com/rest/2.0/ocr/v1/doc_analysis_office"
)
OCR_SUFFIXES = {".pdf", ".jpg", ".jpeg", ".png"}
MAX_FORM_BYTES = 10 * 1024 * 1024
EMPTY_DOCX_MARKDOWN = "[DOCX 中没有可提取的段落或表格文本。]\n"


def get_access_token(api_key: str, secret_key: str) -> str:
    """获取百度 Access Token。"""

    response = requests.post(
        TOKEN_URL,
        params={
            "grant_type": "client_credentials",
            "client_id": api_key,
            "client_secret": secret_key,
        },
        timeout=30,
    )
    response.raise_for_status()
    result = response.json()

    if "access_token" not in result:
        raise RuntimeError(f"获取 Access Token 失败：{result}")

    return result["access_token"]


def _encoded_file(file_bytes: bytes) -> str:
    """将文件内容编码成百度接口要求的 Base64 字符串。"""

    return base64.b64encode(file_bytes).decode("ascii")


def _form_size(data: dict[str, str]) -> int:
    """计算 application/x-www-form-urlencoded 请求体的实际字节数。"""

    return len(urlencode(data).encode("ascii"))


def _request_office_ocr(
    access_token: str,
    file_field: str,
    file_base64: str,
    page_number: int | None = None,
) -> dict[str, Any]:
    """调用一次办公文档识别并返回原始 JSON。"""

    data = {
        file_field: file_base64,
        "language_type": "CHN_ENG",
        "detect_direction": "true",
        "layout_analysis": "true",
        "erase_seal": "true",
    }

    if page_number is not None:
        data["pdf_file_num"] = str(page_number)

    if _form_size(data) > MAX_FORM_BYTES:
        raise ValueError(
            "提交办公文档识别的数据编码后超过 10 MB，"
            "无法满足百度接口限制。"
        )

    response = requests.post(
        OFFICE_OCR_URL,
        params={"access_token": access_token},
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data=data,
        timeout=120,
    )
    response.raise_for_status()
    result = response.json()

    if result.get("error_code"):
        raise RuntimeError(f"办公文档识别失败：{result}")

    if not isinstance(result.get("results"), list):
        raise RuntimeError(f"办公文档识别响应中没有 results：{result}")

    return result


def _referenced_indices(value: Any) -> set[int]:
    """兼容百度版面字段中整数、字符串、列表和对象形式的行号。"""

    if isinstance(value, bool):
        return set()

    if isinstance(value, int):
        return {value}

    if isinstance(value, str):
        return {int(item) for item in re.findall(r"\d+", value)}

    if isinstance(value, list):
        indices: set[int] = set()
        for item in value:
            indices.update(_referenced_indices(item))
        return indices

    if isinstance(value, dict):
        indices = set()
        for key, item in value.items():
            if key in {"idx", "indices", "index"}:
                indices.update(_referenced_indices(item))
        return indices

    return set()


def _result_text(result: Any) -> str:
    """从办公文档识别的单条结果中取出完整文本。"""

    if not isinstance(result, dict):
        return ""

    words = result.get("words")

    if isinstance(words, dict):
        return str(words.get("word") or "").strip()

    if isinstance(words, list):
        parts = []
        for item in words:
            if isinstance(item, dict):
                text = str(item.get("word") or "").strip()
                if text:
                    parts.append(text)
        return " ".join(parts)

    return ""


def office_ocr_result_to_markdown(
    result: dict[str, Any],
    allow_empty: bool = False,
) -> str:
    """将单页办公文档识别结果转换为供 LLM 恢复的中间 Markdown。"""

    results = result.get("results")
    if not isinstance(results, list):
        raise RuntimeError("办公文档识别结果格式无效：results 不是列表。")

    ignored_indices: set[int] = set()
    heading_levels: dict[int, int] = {}

    # 页眉、页脚、页码、脚注和装饰图形不进入简历正文。
    sections = result.get("sections")
    if isinstance(sections, list):
        for section in sections:
            if not isinstance(section, dict):
                continue
            if section.get("attribute") in {
                "header",
                "footer",
                "number",
                "footnote",
            }:
                ignored_indices.update(
                    _referenced_indices(section.get("sec_idx"))
                )

    layouts = result.get("layouts")
    if isinstance(layouts, list):
        for layout in layouts:
            if not isinstance(layout, dict):
                continue

            indices = _referenced_indices(layout.get("layout_idx"))
            layout_type = str(layout.get("layout") or "").lower()

            if layout_type in {"figure", "seal"}:
                ignored_indices.update(indices)
            elif layout_type == "doc_title":
                for index in indices:
                    heading_levels[index] = 1
            elif layout_type in {
                "title",
                "text_title",
                "table_title",
                "figure_title",
            }:
                for index in indices:
                    heading_levels[index] = 2

    lines = []
    for index, item in enumerate(results):
        if index in ignored_indices:
            continue

        text = _result_text(item)
        if not text:
            continue

        heading_level = heading_levels.get(index)
        if heading_level:
            text = f"{'#' * heading_level} {text}"

        lines.append(text)

    if not lines:
        if allow_empty:
            return ""
        raise RuntimeError("办公文档识别成功，但没有返回可用正文。")

    return "\n\n".join(lines).strip() + "\n"


def _render_pdf_page(document: pymupdf.Document, page_index: int) -> str:
    """将超限 PDF 的单页渲染成 JPEG，供办公文档识别按图片处理。"""

    page = document.load_page(page_index)
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5), alpha=False)
    image_base64 = _encoded_file(
        pixmap.tobytes("jpeg", jpg_quality=85)
    )

    probe = {
        "image": image_base64,
        "language_type": "CHN_ENG",
        "detect_direction": "true",
        "layout_analysis": "true",
        "erase_seal": "true",
    }
    if _form_size(probe) > MAX_FORM_BYTES:
        raise ValueError(
            f"PDF 第 {page_index + 1} 页渲染后仍超过百度接口 10 MB 限制。"
        )

    return image_base64


def _recognize_pdf(input_path: Path, access_token: str) -> list[dict[str, Any]]:
    """逐页识别 PDF，必要时先将页面渲染成图片。"""

    document = pymupdf.open(input_path)
    try:
        if document.page_count < 1:
            raise ValueError("PDF 没有可识别的页面。")

        pdf_base64 = _encoded_file(input_path.read_bytes())
        pdf_probe = {
            "pdf_file": pdf_base64,
            "pdf_file_num": str(document.page_count),
            "language_type": "CHN_ENG",
            "detect_direction": "true",
            "layout_analysis": "true",
            "erase_seal": "true",
        }
        upload_pdf_directly = _form_size(pdf_probe) <= MAX_FORM_BYTES
        page_results = []

        for page_number in range(1, document.page_count + 1):
            print(
                f"    正在识别 PDF 第 {page_number}/{document.page_count} 页……"
            )

            if upload_pdf_directly:
                result = _request_office_ocr(
                    access_token=access_token,
                    file_field="pdf_file",
                    file_base64=pdf_base64,
                    page_number=page_number,
                )
            else:
                result = _request_office_ocr(
                    access_token=access_token,
                    file_field="image",
                    file_base64=_render_pdf_page(
                        document,
                        page_number - 1,
                    ),
                )

            page_results.append(result)

        return page_results
    finally:
        document.close()


def _recognize_image(input_path: Path, access_token: str) -> list[dict[str, Any]]:
    """识别单张 JPG、JPEG 或 PNG 图片。"""

    print("    正在识别图片……")
    return [
        _request_office_ocr(
            access_token=access_token,
            file_field="image",
            file_base64=_encoded_file(input_path.read_bytes()),
        )
    ]


def _paragraph_markdown(paragraph: Paragraph) -> str:
    """尽量保留 DOCX 段落的标题或列表语义。"""

    text = paragraph.text.strip()
    if not text:
        return ""

    style_name = str(getattr(paragraph.style, "name", "") or "")
    heading_match = re.search(r"(?:Heading|标题)\s*([1-6])", style_name, re.I)

    if heading_match:
        return f"{'#' * int(heading_match.group(1))} {text}"

    if "Title" in style_name or "标题" == style_name:
        return f"# {text}"

    paragraph_properties = paragraph._p.pPr
    has_numbering = (
        paragraph_properties is not None
        and paragraph_properties.numPr is not None
    )
    if has_numbering or "List" in style_name or "列表" in style_name:
        return f"- {text}"

    return text


def _table_markdown(table: Table) -> str:
    """将 DOCX 表格转换为不丢失单元格内容的 Markdown 文本行。"""

    rows = []
    for row in table.rows:
        cell_texts = []
        seen_cells: set[int] = set()
        for cell in row.cells:
            cell_identity = id(cell._tc)
            if cell_identity in seen_cells:
                continue
            seen_cells.add(cell_identity)
            cell_texts.append(" ".join(cell.text.split()))

        if any(cell_texts):
            rows.append(" | ".join(cell_texts))

    return "\n".join(rows)


def docx_to_markdown(input_path: Path) -> str:
    """在本地按文档顺序提取 DOCX 段落和表格。"""

    document = Document(input_path)
    blocks = []

    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = _paragraph_markdown(Paragraph(child, document))
        elif child.tag == qn("w:tbl"):
            text = _table_markdown(Table(child, document))
        else:
            text = ""

        if text:
            blocks.append(text)

    if not blocks:
        return EMPTY_DOCX_MARKDOWN

    return "\n\n".join(blocks).strip() + "\n"


def BaiduParser(input_file_path: str, output_dir: str) -> str:
    """
    将 PDF、DOCX 或图片转换成中间 Markdown。

    PDF 和图片调用百度办公文档识别；DOCX 使用 python-docx 本地解析。
    成功时返回生成文件的绝对路径。
    """

    input_path = Path(input_file_path)
    if not input_path.exists():
        raise FileNotFoundError(f"找不到输入文件：{input_path.resolve()}")

    suffix = input_path.suffix.lower()
    if suffix not in OCR_SUFFIXES | {".docx"}:
        if suffix == ".doc":
            raise ValueError("不支持旧版 DOC 文件，请先转换为 DOCX 格式。")
        raise ValueError(f"不支持的输入文件格式：{suffix or '无后缀'}")

    output_directory = Path(output_dir)
    output_directory.mkdir(parents=True, exist_ok=True)
    output_path = output_directory / f"{input_path.stem}.md"

    if suffix == ".docx":
        print("    正在本地读取 DOCX 段落和表格……")
        markdown = docx_to_markdown(input_path)
    else:
        load_dotenv(dotenv_path=application_environment_path())
        api_key = os.getenv("BAIDU_API_KEY")
        secret_key = os.getenv("BAIDU_SECRET_KEY")

        if not api_key or not secret_key:
            raise RuntimeError(
                "没有读取到百度密钥，请检查 .env 中是否配置了 "
                "BAIDU_API_KEY 和 BAIDU_SECRET_KEY。"
            )

        print("    正在获取百度 Access Token……")
        access_token = get_access_token(api_key, secret_key)
        print("    正在调用百度办公文档识别（OCR 共享资源包）……")

        if suffix == ".pdf":
            page_results = _recognize_pdf(input_path, access_token)
        else:
            page_results = _recognize_image(input_path, access_token)

        page_markdowns = []
        for page_number, result in enumerate(page_results, start=1):
            page_markdown = office_ocr_result_to_markdown(
                result,
                allow_empty=True,
            )
            if page_markdown:
                page_markdowns.append(page_markdown)
            else:
                print(
                    f"    警告：第 {page_number} 页没有可用正文，已跳过。"
                )

        if not page_markdowns:
            raise RuntimeError("办公文档识别成功，但整份文件没有可用正文。")

        markdown = "\n\n".join(
            item.strip() for item in page_markdowns
        ).strip() + "\n"

    output_path.write_text(markdown, encoding="utf-8")
    print(f"    中间 Markdown 已生成：{output_path.resolve()}")
    return str(output_path.resolve())
