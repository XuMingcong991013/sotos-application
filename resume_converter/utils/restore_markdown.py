"""
调用内网大模型恢复百度解析后的简历Markdown结构。

本模块暴露RestoreMarkdown函数，不包含独立运行入口。
"""

import os
import re
from collections import Counter
from pathlib import Path

import pymupdf
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from dotenv import load_dotenv
from openai import OpenAI

from utils.runtime_paths import application_environment_path


SUPPORTED_SUFFIXES = {
    ".pdf",
    ".docx",
    ".jpg",
    ".jpeg",
    ".png",
}

NO_REFERENCE_TEXT = {
    ".pdf": (
        "[扫描PDF没有可提取的原始文本层；"
        "请仅依据百度Markdown恢复结构。]"
    ),
    ".docx": (
        "[DOCX中没有可提取的段落或表格文本；"
        "请仅依据百度Markdown恢复结构。]"
    ),
    ".jpg": (
        "[当前版本不提取图片辅助参考文本；"
        "请仅依据百度Markdown恢复结构。]"
    ),
    ".jpeg": (
        "[当前版本不提取图片辅助参考文本；"
        "请仅依据百度Markdown恢复结构。]"
    ),
    ".png": (
        "[当前版本不提取图片辅助参考文本；"
        "请仅依据百度Markdown恢复结构。]"
    ),
}


SYSTEM_PROMPT = """
你是一个严格的文档结构恢复程序。

你的任务是修复由PDF、Word或图片解析得到的简历Markdown，使其恢复为结构清晰、阅读顺序正确、忠实于原始文档的Markdown。

你会收到两份数据：

1. baidu_markdown

这是百度文档解析API生成的Markdown，是最终输出内容的主要来源。

它可能存在以下问题：

- 标题层级错误；
- 阅读顺序错误；
- 多栏内容顺序错误；
- 同一条目的左右两部分被拆散；
- 不同文本块被错误合并；
- 原始正文被遗漏；
- PDF换行造成词语断裂；
- 水印或页眉页脚被错误识别为正文。

2. source_reference_text

这是从原始文件中提取的辅助参考文本，主要用于核对：

- 原始物理换行；
- 文本块边界；
- 正文内容遗漏；
- 内容先后顺序；
- 日期、职位、地点等内容的原始位置关系。

原文件参考文本可能为空，也可能包含水印、页眉页脚、重复内容、错乱字符和不正确的跨栏阅读顺序，因此不能直接复制为最终结果。

你可以进行以下操作：

1. 修正Markdown标题层级。
2. 恢复章节、公司、项目、学校和子栏目之间的层级关系。
3. 将被错误移动的日期、职位、地点、学历等内容放回对应条目。
4. 根据上下文恢复多栏文档的正常阅读顺序。
5. 根据原文件参考文本恢复原始物理换行和文本块边界。
6. 拆分被百度错误合并到同一行的不同文本块。
7. 恢复原文件参考文本中存在、但百度Markdown遗漏的正文。
8. 合并由PDF排版或换行造成的明显异常断词。
9. 删除明显不属于正文的水印字符串。
10. 删除重复的页眉和页脚。
11. 删除没有意义的重复空行。
12. 将能够明确识别的列表恢复为Markdown列表。
13. 将能够明确识别的表格保留为Markdown表格或HTML表格。

必须遵守以下规则：

1. 百度Markdown是内容主体，不能随意删除其中的正文。
2. 原文件参考文本中存在、百度Markdown中遗漏的正文可以恢复。
3. 百度Markdown和原文件参考文本中都不存在的内容不得增加。
4. 如果原文件参考文本显示两段内容属于不同物理行或不同文本块，应保留合理分行。
5. 如果同一条目的左右两部分被拆散，只有在归属关系明确时才允许重新组合。
6. 如果内容归属无法确定，应保持百度Markdown原样，不得猜测。
7. 原文件参考文本中的水印、随机字符串、页眉页脚和重复内容不得复制到最终结果。
8. 如果百度Markdown和原文件参考文本发生冲突，优先保留百度Markdown。
9. 只能恢复原文中实际存在的内容，不能根据上下文编造内容。
10. 不得润色、概括、缩写、扩写或改善原文表达。
11. 不得纠正专业术语、英文拼写、语法或事实错误。
12. 不得修改公司、项目、岗位、学校、专业和技术名称。
13. 不得修改电话、邮箱、日期和数字。
14. 最终只输出恢复后的完整Markdown。
15. 不要解释修改过程，不要总结简历。
16. 不要使用Markdown代码块包裹输出结果。

建议使用以下标题层级：

# 候选人姓名或简历标题

## 基本信息
## 优势亮点
## 专业技能
## 工作经历
## 项目经验
## 教育经历
## 获得荣誉
## 语言能力
## 证书及附加信息
## 自我评价

公司、项目和学校条目通常使用三级标题：

### 公司名称｜岗位名称｜时间
### 项目名称｜项目角色｜时间
### 学校名称｜专业｜学历｜时间

项目内部的小标题通常使用四级标题：

#### 项目背景
#### 项目描述
#### 项目功能
#### 项目职责
#### 项目业绩
#### 主控芯片
#### 原理框图
#### PCB设计
#### 个人职责

以上标题只是结构参考。如果原文中不存在某个栏目，不得自行增加。
""".strip()


def extract_pdf_reference_text(
    pdf_path: Path,
) -> str:
    """提取PDF原始文本层，尽量保留物理换行。"""

    document = pymupdf.open(pdf_path)
    page_contents = []
    has_text_layer = False

    try:
        for page_number, page in enumerate(
            document,
            start=1,
        ):
            page_text = page.get_text(
                "text",
                sort=True,
            ).strip()

            if not page_text:
                page_text = (
                    "[本页没有可提取的原始文本层]"
                )
            else:
                has_text_layer = True

            page_contents.append(
                f"===== PDF第{page_number}页 =====\n"
                f"{page_text}"
            )
    finally:
        document.close()

    if not has_text_layer:
        return NO_REFERENCE_TEXT[".pdf"]

    return "\n\n".join(page_contents)


def extract_docx_reference_text(
    docx_path: Path,
) -> str:
    """按文档顺序提取DOCX段落和表格中的文字。"""

    document = Document(docx_path)
    contents = []

    # 只读取段落和表格文字，图片、形状等装饰内容不会进入参考文本。
    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            paragraph_text = Paragraph(
                child,
                document,
            ).text.strip()

            if paragraph_text:
                contents.append(paragraph_text)

        elif child.tag.endswith("}tbl"):
            table = Table(child, document)

            for row in table.rows:
                cell_texts = [
                    " ".join(cell.text.split())
                    for cell in row.cells
                ]

                if any(cell_texts):
                    contents.append(
                        " | ".join(cell_texts)
                    )

    if not contents:
        return NO_REFERENCE_TEXT[".docx"]

    return "\n".join(contents)


def extract_reference_text(
    source_path: Path,
) -> str:
    """按原始文件后缀选择辅助参考文本提取方式。"""

    suffix = source_path.suffix.lower()

    if suffix == ".pdf":
        return extract_pdf_reference_text(source_path)

    if suffix == ".docx":
        return extract_docx_reference_text(source_path)

    if suffix in NO_REFERENCE_TEXT:
        return NO_REFERENCE_TEXT[suffix]

    if suffix == ".doc":
        raise ValueError(
            "不支持旧版DOC文件，请先转换为DOCX格式。"
        )

    raise ValueError(
        f"不支持的原始文件格式：{suffix or '无后缀'}"
    )


def remove_markdown_code_fence(
    text: str,
) -> str:
    """删除模型可能额外添加的Markdown代码块。"""

    text = text.strip()

    if text.startswith("```markdown"):
        text = text[len("```markdown"):].lstrip()
    elif text.startswith("```md"):
        text = text[len("```md"):].lstrip()
    elif text.startswith("```"):
        text = text[3:].lstrip()

    if text.endswith("```"):
        text = text[:-3].rstrip()

    return text.strip() + "\n"


def is_list_marker(
    source_text: str,
    start: int,
    end: int,
) -> bool:
    """
    判断数字是否为列表序号。

    支持以下格式：
    1. 内容
    1、内容
    1) 内容
    """

    line_start = source_text.rfind(
        "\n",
        0,
        start,
    ) + 1

    before_number = source_text[
        line_start:start
    ]

    after_number = source_text[
        end:end + 2
    ]

    return (
        not before_number.strip()
        and bool(
            re.match(
                r"[.、)]\s*",
                after_number,
            )
        )
    )


def extract_important_values(
    text: str,
) -> Counter:
    """
    提取不能被模型修改的重要信息。

    包括：
    - 邮箱；
    - 中国大陆手机号；
    - 日期和日期范围；
    - 百分比；
    - 独立出现的数字。

    不检查列表序号，也不提取水印和随机ID中的数字。
    """

    values = []
    protected_ranges = []

    protected_patterns = [
        # 邮箱
        r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}",

        # 中国大陆手机号
        r"(?<!\d)1[3-9]\d{9}(?!\d)",

        # 日期和日期范围
        (
            r"\d{4}\s*[./年-]\s*\d{1,2}"
            r"(?:\s*[./月-]\s*\d{1,2})?"
            r"(?:\s*(?:-|—|~|至)\s*"
            r"(?:至今|"
            r"\d{4}\s*[./年-]\s*\d{1,2}"
            r"(?:\s*[./月-]\s*\d{1,2})?"
            r"))?"
        ),

        # 百分比
        r"\d+(?:\.\d+)?%",
    ]

    for pattern in protected_patterns:
        for match in re.finditer(
            pattern,
            text,
            flags=re.IGNORECASE,
        ):
            value = re.sub(
                r"\s+",
                "",
                match.group(0),
            )

            values.append(value)

            protected_ranges.append(
                (
                    match.start(),
                    match.end(),
                )
            )

    def overlaps_protected_range(
        start: int,
        end: int,
    ) -> bool:
        """判断数字是否已属于日期、电话等信息。"""

        for protected_start, protected_end in (
            protected_ranges
        ):
            if (
                start < protected_end
                and end > protected_start
            ):
                return True

        return False

    standalone_number_pattern = (
        r"(?<![A-Za-z0-9_])"
        r"\d+(?:\.\d+)?"
        r"(?![A-Za-z0-9_])"
    )

    for match in re.finditer(
        standalone_number_pattern,
        text,
    ):
        if overlaps_protected_range(
            match.start(),
            match.end(),
        ):
            continue

        # 列表序号属于文档结构，不属于简历事实。
        if is_list_marker(
            source_text=text,
            start=match.start(),
            end=match.end(),
        ):
            continue

        values.append(
            match.group(0)
        )

    return Counter(values)


def validate_result(
    original_markdown: str,
    source_reference_text: str,
    restored_markdown: str,
) -> list[str]:
    """
    检查模型是否删除或编造重要信息。

    自动检查只提供警告，不中断文件生成。
    """

    warnings = []

    original_length = len(
        re.sub(
            r"\s+",
            "",
            original_markdown,
        )
    )

    restored_length = len(
        re.sub(
            r"\s+",
            "",
            restored_markdown,
        )
    )

    if restored_length < original_length * 0.80:
        warnings.append(
            "输出内容明显缩短："
            f"百度Markdown约{original_length}字符，"
            f"输出约{restored_length}字符"
        )

    if restored_length > original_length * 1.40:
        warnings.append(
            "输出内容增加较多："
            f"百度Markdown约{original_length}字符，"
            f"输出约{restored_length}字符"
        )

    baidu_values = extract_important_values(
        original_markdown
    )

    source_values = extract_important_values(
        source_reference_text
    )

    restored_values = extract_important_values(
        restored_markdown
    )

    # 百度已经识别出来的重要信息不能丢失。
    missing_values = (
        baidu_values - restored_values
    )

    # 百度Markdown或PDF中存在的信息都允许出现。
    allowed_values = (
        baidu_values | source_values
    )

    # 两个来源都不存在的信息才属于模型新增。
    invented_values = (
        restored_values - allowed_values
    )

    if missing_values:
        missing_text = "、".join(
            f"{value}×{count}"
            for value, count in missing_values.items()
        )

        warnings.append(
            "模型可能删除或修改了百度已识别的"
            f"重要信息：{missing_text}"
        )

    if invented_values:
        invented_text = "、".join(
            f"{value}×{count}"
            for value, count in invented_values.items()
        )

        warnings.append(
            "模型可能新增了百度Markdown和原文件中"
            f"都不存在的重要信息：{invented_text}"
        )

    return warnings


def call_llm_restore(
    original_markdown: str,
    source_reference_text: str,
    base_url: str,
    model: str,
    api_key: str,
) -> str:
    """调用内网大模型恢复Markdown结构。"""

    client = OpenAI(
        api_key=api_key,
        base_url=base_url,
        timeout=300,
    )

    user_prompt = (
        "请严格恢复下面这份简历Markdown的结构。\n\n"
        "百度Markdown是最终内容的主体来源。\n"
        "原文件参考文本只用于核对物理换行、"
        "文本块边界、内容遗漏和阅读顺序。\n"
        "原文件参考文本可能为空，或包含水印、页眉页脚、"
        "重复内容或错乱字符，不得直接照搬。\n\n"
        "<baidu_markdown>\n"
        f"{original_markdown}\n"
        "</baidu_markdown>\n\n"
        "<source_reference_text>\n"
        f"{source_reference_text}\n"
        "</source_reference_text>"
    )

    response = client.chat.completions.create(
        model=model,
        temperature=0,
        max_tokens=16000,
        messages=[
            {
                "role": "system",
                "content": SYSTEM_PROMPT,
            },
            {
                "role": "user",
                "content": user_prompt,
            },
        ],
    )

    restored_markdown = (
        response.choices[0].message.content
    )

    if not restored_markdown:
        raise RuntimeError(
            "模型没有返回内容。"
        )

    return remove_markdown_code_fence(
        restored_markdown
    )


def RestoreMarkdown(
    input_file_path: str,
    input_file: str,
    output_file: str,
) -> str:
    """
    恢复百度Markdown结构。

    Args:
        input_file_path:
            原始PDF、Word或图片文件路径。

        input_file:
            百度解析生成的Markdown文件路径。

        output_file:
            恢复结构后的Markdown输出路径。

    Returns:
        最终Markdown文件的绝对路径。
    """

    load_dotenv(dotenv_path=application_environment_path())

    source_path = Path(input_file_path)
    input_path = Path(input_file)
    output_path = Path(output_file)

    if not source_path.exists():
        raise FileNotFoundError(
            f"找不到原始文件：{source_path.resolve()}"
        )

    if not input_path.exists():
        raise FileNotFoundError(
            "找不到百度Markdown文件："
            f"{input_path.resolve()}"
        )

    if source_path.suffix.lower() not in SUPPORTED_SUFFIXES:
        if source_path.suffix.lower() == ".doc":
            raise ValueError(
                "不支持旧版DOC文件，请先转换为DOCX格式。"
            )

        raise ValueError(
            "不支持的原始文件格式："
            f"{source_path.suffix.lower() or '无后缀'}"
        )

    if input_path.suffix.lower() != ".md":
        raise ValueError(
            "input_file必须指向Markdown文件。"
        )

    api_key = os.getenv("LLM_API_KEY")
    base_url = os.getenv("LLM_BASE_URL")
    model = os.getenv("LLM_MODEL")

    if not api_key:
        raise RuntimeError(
            "没有读取到LLM_API_KEY，"
            "请检查.env文件。"
        )

    if not base_url:
        raise RuntimeError(
            "没有读取到LLM_BASE_URL，"
            "请检查.env文件。"
        )

    if not model:
        raise RuntimeError(
            "没有读取到LLM_MODEL，"
            "请检查.env文件。"
        )

    print(f"    原始文件：{source_path.resolve()}")
    print(f"    百度 Markdown：{input_path.resolve()}")
    print(f"    使用模型：{model}")

    print("    正在读取百度 Markdown……")

    original_markdown = input_path.read_text(
        encoding="utf-8-sig"
    )

    if not original_markdown.strip():
        raise RuntimeError(
            "百度Markdown文件为空。"
        )

    print("    正在提取原文件辅助参考文本……")

    source_reference_text = (
        extract_reference_text(
            source_path
        )
    )

    print("    正在调用 LLM 恢复 Markdown 结构……")

    restored_markdown = call_llm_restore(
        original_markdown=original_markdown,
        source_reference_text=source_reference_text,
        base_url=base_url,
        model=model,
        api_key=api_key,
    )

    print("    正在检查关键数据是否被修改……")

    validation_warnings = validate_result(
        original_markdown=original_markdown,
        source_reference_text=source_reference_text,
        restored_markdown=restored_markdown,
    )

    if validation_warnings:
        print(
            "\n警告：结构恢复结果没有完全"
            "通过自动检查："
        )

        for warning in validation_warnings:
            print(f"      - {warning}")

        print(
            "自动检查只作为提示，"
            "不会中断文件生成。"
        )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path.write_text(
        restored_markdown,
        encoding="utf-8",
    )

    print("    Markdown 结构恢复完成。")
    print(f"    输出文件：{output_path.resolve()}")

    return str(output_path.resolve())
