"""
调用内网大模型提取简历信息并生成可追溯的过程数据。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from collections.abc import Callable
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
- 项目描述是承载项目原文的完整内容容器，不是摘要字段。项目标题下属于该项目的“项目描述”“应用技术”“主要职责”“功能说明”“实现细节”“成果”等正文、子标题和列表都必须按原有Markdown顺序完整放入description，不得只摘取其中一段。
- 原文项目内容不能安全归入其他字段时，也必须原样保留在description中；固定字段结构不能成为删除原文的理由。
- 项目描述必须逐字来自原文，不得概括、改写、合并相似条目或省略技术列表。即使没有独立项目栏目，也不能仅凭这一点返回空的project_experiences。
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

CHERY_SYSTEM_PROMPT = """
你是一个严格、可追溯的奇瑞标准简历信息提取程序。

除“自我评价”和“专业技能”的明确兜底规则外，你只能提取输入Markdown中确实存在的信息。宁可留空，也不能根据姓名、照片、岗位、公司业务、学校层级、行业常识或上下文猜测。

需要提取：
1. 基本信息：姓名、性别、出生年份、出生年月、联系电话、邮箱、籍贯。出生年份是出生年月中的年份子串；其他字段不得相互推断。
2. 自我评价：有独立原文时逐字提取并标记explicit；没有时，必须仅依据简历中可核验的工作、项目、教育和技能事实生成一段简短评价，标记generated并列出逐字原文依据。不得加入原文没有的年限、性格、能力、成果或程度判断。只有整份简历不存在任何可作为依据的有效事实时才允许返回missing。
3. 专业技能：有明确栏目时保留原文Markdown并标记explicit；没有明确栏目时，必须仅依据可核验的工作、项目或教育事实生成一段简短技能说明，标记generated并列出逐字原文依据，不得只把零散原文片段拼接成inferred内容。只有整份简历不存在任何可作为依据的有效事实时才允许返回missing。
4. 工作经历：原始时间、公司、岗位。工作描述可继续提取供过程数据使用。
5. 项目经历：项目名称、项目描述、工作业绩；岗位和时间可继续提取供过程数据使用。项目描述和工作业绩必须分别来自原文，不能把同一句内容无依据地重复放入两个字段。
6. 教育经历：原始时间、学校、学历、专业。

项目识别规则：
- 项目不要求出现在独立的项目栏目中；工作经历下明确命名的项目、产品、系统、平台、车型或设备开发也属于项目候选。
- 边界明确时分别提取；多个名称共用描述且无法安全归属时按原文产品组整体保留，不能擅自拆分。
- 项目时间只有在项目明确隶属于某段工作时才能沿用；项目岗位没有明确原文时留空。
- 工作业绩必须是原文明确写出的交付结果、量化结果、奖项、改善效果或已完成成果；只有职责描述时必须留空。
- “项目名称、项目描述、工作业绩”只是输出结构，不是内容筛选条件。每个项目标题下的全部原文，包括“应用技术”“主要职责”“功能说明”“实现细节”等子标题和列表，都必须按原有Markdown顺序完整保留。
- 明确属于工作业绩的原文放入achievement；其余无法归入固定字段的项目原文一律放入description。不得因achievement为空或内容不符合三字段结构而删除任何项目正文。
- description不得概括、改写或只保留项目简介；必须逐字包含该项目除已单独提取工作业绩外的全部原文内容。

证据与生成规则：
- 所有事实字段和explicit/inferred文本必须能在输入Markdown中逐字核验。
- generated只允许用于自我评价和专业技能，且evidence中的每一条都必须逐字来自输入Markdown。
- generated正文不要自行添加“AI生成”等提示，程序会统一添加醒目标记。
- 原文缺失返回空字符串并在issues中说明missing；含糊、冲突或无法归属时返回空字符串并说明ambiguous。
- 现居地不能作为籍贯；姓名不能推断性别；年龄不能反推出生年月；只写研究生不能判断硕士或博士。
- 不计算工作年限、不补工作或项目月份、不排序；这些由程序处理。

只输出一个JSON对象，不要使用Markdown代码块，不要解释：
{
  "basic_information": {
    "name": {"value": "", "evidence": ""},
    "gender": {"value": "", "evidence": ""},
    "birth_year": {"value": "", "evidence": ""},
    "birth_date": {"value": "", "evidence": ""},
    "phone": {"value": "", "evidence": ""},
    "email": {"value": "", "evidence": ""},
    "native_place": {"value": "", "evidence": ""}
  },
  "self_evaluation": {
    "text": "",
    "source_type": "explicit或generated或missing",
    "evidence": ["逐字原文片段"],
    "note": ""
  },
  "professional_skills": {
    "text": "",
    "source_type": "explicit或generated或missing",
    "evidence": ["逐字原文片段"],
    "note": ""
  },
  "work_experiences": [
    {
      "original_time": "", "company_name": "", "position_name": "", "description": "",
      "evidence": {"original_time": "", "company_name": "", "position_name": "", "description": ""}
    }
  ],
  "project_experiences": [
    {
      "project_name": "", "position_name": "", "original_time": "", "description": "", "achievement": "",
      "evidence": {"project_name": "", "position_name": "", "original_time": "", "description": "", "achievement": ""}
    }
  ],
  "education_experiences": [
    {
      "original_time": "", "school_name": "", "degree": "", "major": "",
      "evidence": {"original_time": "", "school_name": "", "degree": "", "major": ""}
    }
  ],
  "issues": [
    {"field": "字段路径", "status": "missing或ambiguous", "evidence": "", "note": ""}
  ]
}
""".strip()

CHERY_NARRATIVE_FALLBACK_PROMPT = """
你是奇瑞标准简历的受控内容补全程序。输入是一份恢复后的简历Markdown。

只处理“自我评价”和“专业技能”两个字段。调用方已经确认首轮提取未能生成这两个字段中的至少一个，因此：
- 对缺失的自我评价，必须根据简历中逐字可核验的工作、项目、教育或技能事实生成一段简短文字。
- 对缺失的专业技能，必须根据简历中逐字可核验的工作、项目或教育事实生成一段简短文字。
- 不能添加原文没有的年限、性格、能力程度、技术、成果、证书或评价。
- 每个generated字段必须列出至少一条逐字存在于输入Markdown中的证据；证据应优先选择完整、简短的原文句子或片段。
- 只有整份简历不存在任何可作为依据的有效事实时，才允许返回missing。
- 不要在正文中添加“AI生成”等提示，程序会统一标记。

只输出一个JSON对象，不要使用Markdown代码块，不要解释：
{
  "self_evaluation": {
    "text": "",
    "source_type": "generated或missing",
    "evidence": ["逐字原文片段"],
    "note": ""
  },
  "professional_skills": {
    "text": "",
    "source_type": "generated或missing",
    "evidence": ["逐字原文片段"],
    "note": ""
  }
}
""".strip()


def call_llm_extract(
    restored_markdown: str,
    base_url: str,
    model: str,
    api_key: str,
) -> dict[str, Any]:
    """调用内网模型，并在JSON格式错误时仅修复格式一次。"""

    return _call_llm_extract(
        restored_markdown,
        base_url,
        model,
        api_key,
        SYSTEM_PROMPT,
    )


def call_llm_extract_chery(
    restored_markdown: str,
    base_url: str,
    model: str,
    api_key: str,
) -> dict[str, Any]:
    """使用奇瑞专用字段和受控生成规则调用内网模型。"""

    return _call_llm_extract(
        restored_markdown,
        base_url,
        model,
        api_key,
        CHERY_SYSTEM_PROMPT,
    )


def _call_llm_complete_chery_narratives(
    restored_markdown: str,
    base_url: str,
    model: str,
    api_key: str,
) -> dict[str, Any]:
    """对奇瑞首轮漏生成的两个长文本字段执行一次受控补全。"""

    return _call_llm_extract(
        restored_markdown,
        base_url,
        model,
        api_key,
        CHERY_NARRATIVE_FALLBACK_PROMPT,
    )


def _call_llm_extract(
    restored_markdown: str,
    base_url: str,
    model: str,
    api_key: str,
    system_prompt: str,
) -> dict[str, Any]:
    """按给定模板提示调用模型，并在JSON错误时修复格式一次。"""

    client = OpenAI(
        api_key=api_key,
        base_url=base_url,
        timeout=300,
    )
    messages = [
        {"role": "system", "content": system_prompt},
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


ExtractionCaller = Callable[[str, str, str, str], dict[str, Any]]


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

    return _information_extractor(
        input_file_path=input_file_path,
        input_file=input_file,
        output_dir=output_dir,
        extraction_caller=call_llm_extract,
        template_tag="SOTOS",
    )


def CheryInformationExtractor(
    input_file_path: str,
    input_file: str,
    output_dir: str,
) -> str | None:
    """提取奇瑞模板所需字段，并保留AI兜底内容的来源标记。"""

    return _information_extractor(
        input_file_path=input_file_path,
        input_file=input_file,
        output_dir=output_dir,
        extraction_caller=call_llm_extract_chery,
        template_tag="奇瑞",
    )


def _information_extractor(
    input_file_path: str,
    input_file: str,
    output_dir: str,
    extraction_caller: ExtractionCaller,
    template_tag: str,
) -> str | None:
    """执行共用的提取调用、校验、保存、汇总和错误处理。"""

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
        raw_data = extraction_caller(
            restored_markdown=restored_markdown,
            base_url=base_url,
            model=model,
            api_key=api_key,
        )
        normalized = normalize_extraction(
            raw_data=raw_data,
            source_markdown=restored_markdown,
            template_tag=template_tag,
        )
        if template_tag == "奇瑞" and _missing_chery_narratives(normalized):
            print("  正在补全奇瑞模板的自我评价或专业技能……")
            fallback_data = _call_llm_complete_chery_narratives(
                restored_markdown=restored_markdown,
                base_url=base_url,
                model=model,
                api_key=api_key,
            )
            raw_data = _merge_chery_narrative_fallback(
                raw_data,
                fallback_data,
                normalized,
            )
            normalized = normalize_extraction(
                raw_data=raw_data,
                source_markdown=restored_markdown,
                template_tag=template_tag,
            )
        metadata = {
            "schema_version": "1.1" if template_tag == "奇瑞" else "1.0",
            "candidate_id": _candidate_id(source_path),
            "source_file": str(source_path.resolve()),
            "restored_markdown_file": str(restored_path.resolve()),
            "extracted_at": datetime.now().isoformat(timespec="seconds"),
        }
        if template_tag == "奇瑞":
            metadata["template_tag"] = template_tag
        normalized.update(metadata)
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


def _missing_chery_narratives(data: dict[str, Any]) -> bool:
    """判断奇瑞两个允许生成的长文本字段是否仍为空。"""

    return any(
        not _text(data.get(field))
        for field in ("self_evaluation", "professional_skills")
    )


def _merge_chery_narrative_fallback(
    raw_data: dict[str, Any],
    fallback_data: dict[str, Any],
    normalized: dict[str, Any],
) -> dict[str, Any]:
    """只回填首轮缺失字段，并清理已被补全的旧缺失提示。"""

    merged = dict(raw_data)
    replaced_fields: set[str] = set()

    for field in ("self_evaluation", "professional_skills"):
        if _text(normalized.get(field)):
            continue

        candidate = fallback_data.get(field)
        if not isinstance(candidate, dict):
            continue

        merged[field] = candidate
        if _text(candidate.get("text")):
            replaced_fields.add(field)

    if replaced_fields:
        merged["issues"] = [
            issue
            for issue in merged.get("issues", [])
            if not (
                isinstance(issue, dict)
                and _text(issue.get("field")) in replaced_fields
            )
        ]

    return merged


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


def _text(value: Any) -> str:
    """把模型字段安全转换为去除首尾空白的文本。"""

    return value.strip() if isinstance(value, str) else ""
