"""
调用内网大模型提取简历信息并生成可追溯的过程数据。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI

from utils.error_logging import write_error_log
from utils.processing_records import (
    TRACKING_WORKBOOK_NAME,
    update_extraction_record,
)
from utils.runtime_paths import application_environment_path

from .normalizer import normalize_extraction
from .storage import (
    rebuild_summary_workbook,
    save_extracted_json,
)


INFORMATION_EXTRACTION_DIRNAME = "information_extraction"

SYSTEM_PROMPT = """
你是一个严格、保守的简历事实提取程序。

你只能提取输入Markdown中确实存在的信息。宁可留空，也不能根据姓名、照片、岗位、公司业务、学校层级、行业常识或上下文猜测。

需要提取：
1. 基本资料：姓名、性别、出生年份、籍贯。
2. 专业技能：作为一段Markdown文本保留原文换行、序号和项目符号。若没有独立技能栏目，只能返回工作或项目中明确体现技能的原文片段，不得改写或概括。
3. 工作经历：原始时间、公司、岗位、工作描述。
4. 项目经历：项目名称、岗位或角色、原始时间、项目描述。
5. 教育经历：原始时间、学校、学历、专业。

项目识别规则：
- 项目不要求出现在独立的“项目经历”或“项目经验”栏目中。工作经历下明确写出的项目、产品、系统、平台、车型或设备开发，也必须作为项目候选提取。
- “主要产品有”“负责/参与……开发”“……系统”“……平台”“……项目”等原文内容都可能是项目证据；同一段原文可以同时作为工作描述和项目描述的证据。
- 若同一工作时间段列出多个项目或产品，且名称与描述的边界明确，应分别提取；若多个名称共用一段描述、无法安全确定逐项归属，则按原文产品组整体保留为一个项目，不能擅自拆分或编造对应关系。
- 项目时间只有在项目内容明确隶属于某一工作时间段时，才可使用该段原始时间；项目岗位或角色没有明确写出时必须留空。
- 项目描述必须逐字来自原文。即使没有独立项目栏目，也不能仅凭这一点返回空的project_experiences。
- 不能把公司主营业务、岗位常识或仅有的技能词推断为项目。

证据规则：
- 每个非空事实必须提供可以从输入Markdown中逐字找到的证据。
- description证据必须是原文描述，不得生成新描述。
- 原文没有的信息返回空字符串，并在issues中说明missing。
- 原文含糊、无法确定归属或存在冲突时返回空字符串，并在issues中说明ambiguous。
- 现居地不能作为籍贯；姓名不能用于推断性别；年龄不能反推出生年份。
- 只写“研究生”时不能擅自判断硕士或博士。
- 不计算工作年限，不补充任何月份，不排序；这些由程序完成。

只输出一个JSON对象，不要使用Markdown代码块，不要解释。必须使用以下结构：
{
  "basic_information": {
    "name": {"value": "", "evidence": ""},
    "gender": {"value": "", "evidence": ""},
    "birth_year": {"value": "", "evidence": ""},
    "native_place": {"value": "", "evidence": ""}
  },
  "professional_skills": {
    "text": "",
    "source_type": "explicit或inferred或missing",
    "evidence": ["逐字原文片段"],
    "note": ""
  },
  "work_experiences": [
    {
      "original_time": "",
      "company_name": "",
      "position_name": "",
      "description": "",
      "evidence": {
        "original_time": "",
        "company_name": "",
        "position_name": "",
        "description": ""
      }
    }
  ],
  "project_experiences": [
    {
      "project_name": "",
      "position_name": "",
      "original_time": "",
      "description": "",
      "evidence": {
        "project_name": "",
        "position_name": "",
        "original_time": "",
        "description": ""
      }
    }
  ],
  "education_experiences": [
    {
      "original_time": "",
      "school_name": "",
      "degree": "",
      "major": "",
      "evidence": {
        "original_time": "",
        "school_name": "",
        "degree": "",
        "major": ""
      }
    }
  ],
  "issues": [
    {"field": "字段路径", "status": "missing或ambiguous", "evidence": "", "note": ""}
  ]
}
""".strip()


def call_llm_extract(
    restored_markdown: str,
    base_url: str,
    model: str,
    api_key: str,
) -> dict[str, Any]:
    """调用内网模型，并在JSON格式错误时仅修复格式一次。"""

    client = OpenAI(
        api_key=api_key,
        base_url=base_url,
        timeout=300,
    )
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                "请从以下恢复后的简历Markdown中提取信息：\n\n"
                "<restored_markdown>\n"
                f"{restored_markdown}\n"
                "</restored_markdown>"
            ),
        },
    ]
    response = client.chat.completions.create(
        model=model,
        temperature=0,
        max_tokens=16000,
        messages=messages,
    )
    content = _response_content(response)

    try:
        return _parse_json_object(content)
    except (json.JSONDecodeError, ValueError):
        repair_response = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=16000,
            messages=[
                *messages,
                {"role": "assistant", "content": content},
                {
                    "role": "user",
                    "content": (
                        "上一个响应不是有效JSON。只能修复JSON语法和结构，"
                        "不得增加、删除或修改任何事实内容。只返回有效JSON对象。"
                    ),
                },
            ],
        )
        return _parse_json_object(
            _response_content(repair_response)
        )


def InformationExtractor(
    input_file_path: str,
    input_file: str,
    output_dir: str,
) -> str | None:
    """
    从恢复后的Markdown提取简历信息。

    Args:
        input_file_path: 原始简历文件路径。
        input_file: RestoreMarkdown生成的Markdown路径。
        output_dir: 整个处理任务的输出根目录。

    Returns:
        单份候选人结构化JSON的绝对路径。

        失败时返回None，并把异常写入过程数据错误日志。
    """

    load_dotenv(dotenv_path=application_environment_path())
    source_path = Path(input_file_path)
    restored_path = Path(input_file)
    output_root = Path(output_dir)
    process_directory = output_root / "process_data"
    extraction_directory = (
        process_directory / INFORMATION_EXTRACTION_DIRNAME
    )
    tracking_workbook = (
        process_directory / TRACKING_WORKBOOK_NAME
    )
    json_path = (
        extraction_directory
        / f"{source_path.stem}_extracted.json"
    )

    try:
        if not source_path.exists():
            raise FileNotFoundError(
                f"找不到原始简历文件：{source_path.resolve()}"
            )

        if not restored_path.exists():
            raise FileNotFoundError(
                "找不到结构恢复Markdown："
                f"{restored_path.resolve()}"
            )

        if restored_path.suffix.lower() != ".md":
            raise ValueError("input_file必须指向Markdown文件。")

        api_key = os.getenv("LLM_API_KEY")
        base_url = os.getenv("LLM_BASE_URL")
        model = os.getenv("LLM_MODEL")

        if not api_key or not base_url or not model:
            raise RuntimeError(
                "LLM配置不完整，请检查.env中的LLM_API_KEY、"
                "LLM_BASE_URL和LLM_MODEL。"
            )

        restored_markdown = restored_path.read_text(
            encoding="utf-8-sig"
        )

        if not restored_markdown.strip():
            raise RuntimeError("结构恢复Markdown文件为空。")

        print("  正在调用 LLM 提取简历信息……")
        raw_data = call_llm_extract(
            restored_markdown=restored_markdown,
            base_url=base_url,
            model=model,
            api_key=api_key,
        )
        normalized = normalize_extraction(
            raw_data=raw_data,
            source_markdown=restored_markdown,
        )
        normalized.update(
            {
                "schema_version": "1.0",
                "candidate_id": _candidate_id(source_path),
                "source_file": str(source_path.resolve()),
                "restored_markdown_file": str(
                    restored_path.resolve()
                ),
                "extracted_at": datetime.now().isoformat(
                    timespec="seconds"
                ),
            }
        )
        result_path = save_extracted_json(
            normalized,
            json_path,
        )
        rebuild_summary_workbook(extraction_directory)

        if tracking_workbook.exists():
            update_extraction_record(
                workbook_path=tracking_workbook,
                source_path=source_path,
                restored_markdown_path=restored_path,
                json_path=Path(result_path),
                status="成功",
            )

        print(f"  [信息提取完成] {result_path}")
        return str(result_path)

    except Exception as error:
        error_file = write_error_log(
            output_dir=process_directory,
            input_file_path=source_path,
            stage="信息提取",
        )

        if tracking_workbook.exists():
            try:
                update_extraction_record(
                    workbook_path=tracking_workbook,
                    source_path=source_path,
                    restored_markdown_path=restored_path,
                    json_path=(
                        json_path
                        if json_path.exists()
                        else None
                    ),
                    status="失败",
                    error_message=str(error),
                )
            except Exception:
                pass

        print(f"  [信息提取失败] {error}")
        print(f"  详细错误已写入：{error_file.resolve()}")
        return None


def _candidate_id(source_path: Path) -> str:
    """使用文件名和绝对路径摘要生成稳定候选人ID。"""

    digest = hashlib.sha256(
        str(source_path.resolve()).encode("utf-8")
    ).hexdigest()[:10]
    return f"{source_path.stem}_{digest}"


def _response_content(response) -> str:
    """读取模型响应正文并检查空结果。"""

    try:
        content = response.choices[0].message.content
    except (AttributeError, IndexError) as error:
        raise RuntimeError("模型响应结构无效。") from error

    if not content:
        raise RuntimeError("模型没有返回信息提取内容。")

    return content.strip()


def _parse_json_object(content: str) -> dict[str, Any]:
    """移除模型可能添加的代码围栏并解析JSON对象。"""

    content = re.sub(
        r"^```(?:json)?\s*|\s*```$",
        "",
        content.strip(),
        flags=re.IGNORECASE,
    )
    data = json.loads(content)

    if not isinstance(data, dict):
        raise ValueError("模型返回的JSON顶层必须是对象。")

    return data
