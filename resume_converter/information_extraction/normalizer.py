"""
简历提取结果的证据校验、日期标准化与工作年限计算。
"""

from __future__ import annotations

import math
import re
from datetime import date
from typing import Any


BASIC_FIELDS = {
    "name": "姓名",
    "gender": "性别",
    "birth_year": "出生年份",
    "native_place": "籍贯",
}

CHERY_BASIC_FIELDS = {
    "birth_date": "出生年月",
    "phone": "联系电话",
    "email": "邮箱",
}

EMAIL_REVIEW_MARKER = "<!-- EMAIL_REQUIRES_MANUAL_REVIEW -->"


def normalize_extraction(
    raw_data: dict[str, Any],
    source_markdown: str,
    calculation_date: date | None = None,
    template_tag: str = "SOTOS",
) -> dict[str, Any]:
    """将LLM结果转换为可供Word模板直接使用的可信数据。"""

    is_chery = template_tag.strip().upper() == "奇瑞"
    issues: list[dict[str, str]] = []
    field_evidence: dict[str, Any] = {}
    basic_information: dict[str, str] = {}
    raw_basic = _mapping(raw_data.get("basic_information"))
    email_requires_manual_review = (
        EMAIL_REVIEW_MARKER in source_markdown
    )

    basic_fields = dict(BASIC_FIELDS)
    if is_chery:
        basic_fields.update(CHERY_BASIC_FIELDS)

    for field, label in basic_fields.items():
        item = _field_item(raw_basic.get(field))
        value = item["value"]
        evidence = item["evidence"]

        evidence_valid = _value_has_evidence(
            value,
            evidence,
            source_markdown,
        )

        if (
            field == "native_place"
            and evidence_valid
            and not re.search(r"籍贯|祖籍", evidence)
        ):
            evidence_valid = False

        if field == "email" and email_requires_manual_review:
            evidence_valid = False

        if value and evidence_valid:
            basic_information[field] = value
            field_evidence[
                f"basic_information.{field}"
            ] = evidence
        else:
            basic_information[field] = ""
            _append_issue(
                issues,
                f"basic_information.{field}",
                "ambiguous" if evidence else "missing",
                evidence,
                (
                    f"简历中没有可核验的{label}信息"
                    if not evidence
                    else (
                        "原文存在多个相似邮箱候选，需要人工确认"
                        if field == "email" and email_requires_manual_review
                        else (
                            "原文没有明确标注籍贯或祖籍，不能用现居地等信息代替"
                            if field == "native_place"
                            else f"{label}与原文依据无法相互验证"
                        )
                    )
                ),
            )

    skills, skills_source, skills_evidence = (
        _normalize_skills(
            raw_data.get("professional_skills"),
            source_markdown,
            issues,
            allow_inferred=not is_chery,
            allow_generated=is_chery,
        )
    )
    field_evidence["professional_skills"] = (
        skills_evidence
    )

    work_experiences = _normalize_experience_list(
        raw_data.get("work_experiences"),
        "work_experiences",
        source_markdown,
        issues,
        allow_partial=is_chery,
    )
    work_experiences.sort(
        key=lambda item: _date_sort_key(
            item.get("start_date", "")
        ),
        reverse=True,
    )

    project_experiences = _normalize_experience_list(
        raw_data.get("project_experiences"),
        "project_experiences",
        source_markdown,
        issues,
        include_achievement=is_chery,
        allow_partial=is_chery,
    )
    project_experiences = _restore_explicit_project_blocks(
        project_experiences,
        source_markdown,
        issues,
        include_achievement=is_chery,
    )

    education_experiences = _normalize_education_list(
        raw_data.get("education_experiences"),
        source_markdown,
        issues,
    )
    education_experiences.sort(
        key=lambda item: _date_sort_key(
            item.get("end_date", "")
        ),
        reverse=True,
    )

    work_years = calculate_work_years(
        education_experiences,
        calculation_date=calculation_date,
    )

    if work_years["value"] is None:
        _append_issue(
            issues,
            "work_years",
            work_years["status"],
            work_years["evidence"],
            work_years["note"],
        )

    for raw_issue in _list(raw_data.get("issues")):
        issue = _mapping(raw_issue)
        _append_issue(
            issues,
            _text(issue.get("field")) or "unknown",
            _text(issue.get("status")) or "ambiguous",
            _text(issue.get("evidence")),
            _text(issue.get("note")) or "模型标记该信息不明确",
        )

    issues = _remove_resolved_project_issues(
        issues,
        project_experiences,
    )

    result = {
        "basic_information": basic_information,
        "professional_skills": skills,
        "professional_skills_source": skills_source,
        "work_experiences": work_experiences,
        "project_experiences": project_experiences,
        "education_experiences": education_experiences,
        "work_years": work_years,
        "field_evidence": field_evidence,
        "extraction_issues": _deduplicate_issues(issues),
    }

    if is_chery:
        evaluation, evaluation_source, evaluation_evidence = (
            _normalize_narrative(
                raw_data.get("self_evaluation"),
                source_markdown,
                issues,
                field="self_evaluation",
                label="自我评价",
                allow_inferred=False,
                allow_generated=True,
            )
        )
        result["self_evaluation"] = evaluation
        result["self_evaluation_source"] = evaluation_source
        result["field_evidence"]["self_evaluation"] = (
            evaluation_evidence
        )
        result["extraction_issues"] = _deduplicate_issues(issues)

    return result


def calculate_work_years(
    education_experiences: list[dict[str, Any]],
    calculation_date: date | None = None,
) -> dict[str, Any]:
    """按时间上最后一段教育的毕业年月计算工作年限。"""

    today = calculation_date or date.today()

    if not education_experiences:
        return {
            "value": None,
            "status": "missing",
            "graduation_date": "",
            "calculation_date": f"{today:%Y/%m}",
            "evidence": "",
            "note": "没有可用于计算工作年限的教育经历",
        }

    latest = max(
        education_experiences,
        key=lambda item: _date_sort_key(
            item.get("end_date", "")
        ),
    )
    end_date = _text(latest.get("end_date"))
    evidence = _text(latest.get("original_time"))

    if end_date == "至今":
        return {
            "value": 0.0,
            "status": "calculated",
            "graduation_date": "至今",
            "calculation_date": f"{today:%Y/%m}",
            "evidence": evidence,
            "note": "最后一段教育尚未结束，工作年限按0年计算",
        }

    match = re.fullmatch(r"(\d{4})/(\d{2})", end_date)

    if not match:
        return {
            "value": None,
            "status": "ambiguous",
            "graduation_date": end_date,
            "calculation_date": f"{today:%Y/%m}",
            "evidence": evidence,
            "note": "最后一段教育缺少可确定的毕业年月",
        }

    graduation_year = int(match.group(1))
    graduation_month = int(match.group(2))
    elapsed_months = (
        (today.year - graduation_year) * 12
        + today.month
        - graduation_month
    )

    if elapsed_months < 0:
        return {
            "value": 0.0,
            "status": "calculated",
            "graduation_date": end_date,
            "calculation_date": f"{today:%Y/%m}",
            "evidence": evidence,
            "note": "毕业时间晚于计算时间，工作年限按0年计算",
        }

    value = math.floor(elapsed_months / 6) / 2
    return {
        "value": value,
        "status": "calculated",
        "graduation_date": end_date,
        "calculation_date": f"{today:%Y/%m}",
        "evidence": evidence,
        "note": "按最后一段教育毕业年月计算，并向下取整到0.5年",
    }


def normalize_date_range(
    original_time: str,
    education: bool = False,
) -> dict[str, str]:
    """标准化时间范围；仅教育年份可按9月入学、6月毕业补月。"""

    original_time = _text(original_time)

    if not original_time:
        return {
            "start_date": "",
            "end_date": "",
            "time": "",
            "date_status": "missing",
            "normalization_note": "原文没有时间信息",
        }

    tokens = _extract_date_tokens(original_time)
    has_present = bool(re.search(r"至今|现在|目前", original_time))
    start_date = tokens[0] if tokens else ""
    end_date = "至今" if has_present else (
        tokens[1] if len(tokens) > 1 else ""
    )
    note = ""
    status = "extracted"

    if education and start_date and end_date:
        if re.fullmatch(r"\d{4}", start_date):
            start_date = f"{start_date}/09"
            status = "normalized"
            note = "教育开始年份按9月入学补充月份"

        if re.fullmatch(r"\d{4}", end_date):
            end_date = f"{end_date}/06"
            status = "normalized"
            note = (
                f"{note}；教育结束年份按6月毕业补充月份"
                if note
                else "教育结束年份按6月毕业补充月份"
            )

    display_time = start_date

    if start_date and end_date:
        display_time = f"{start_date}-{end_date}"

    if not start_date:
        status = "ambiguous"
        note = "无法从原文确定开始时间"

    return {
        "start_date": start_date,
        "end_date": end_date,
        "time": display_time or original_time,
        "date_status": status,
        "normalization_note": note,
    }


def _normalize_skills(
    raw_value: Any,
    source_markdown: str,
    issues: list[dict[str, str]],
    allow_inferred: bool = True,
    allow_generated: bool = False,
) -> tuple[str, str, list[str]]:
    """保留技能原文形态，推导型技能只拼接可核验原文片段。"""

    return _normalize_narrative(
        raw_value,
        source_markdown,
        issues,
        field="professional_skills",
        label="专业技能",
        allow_inferred=allow_inferred,
        allow_generated=allow_generated,
    )


def _normalize_narrative(
    raw_value: Any,
    source_markdown: str,
    issues: list[dict[str, str]],
    field: str,
    label: str,
    allow_inferred: bool,
    allow_generated: bool,
) -> tuple[str, str, list[str]]:
    """校验原文型或奇瑞专用AI生成型长文本及其依据。"""

    item = _mapping(raw_value)
    text = _text(item.get("text"))
    source_type = _text(item.get("source_type")).lower()
    evidence = [
        _text(value)
        for value in _list(item.get("evidence"))
        if _text(value)
    ]
    valid_evidence = [
        value
        for value in evidence
        if _supported(value, source_markdown)
    ]

    allowed_sources = {"explicit"}
    if allow_inferred:
        allowed_sources.add("inferred")
    if allow_generated:
        allowed_sources.add("generated")

    if source_type not in allowed_sources:
        source_type = "missing"

    if source_type == "explicit" and _supported(
        text,
        source_markdown,
    ):
        return text, "explicit", valid_evidence or [text]

    if source_type == "generated" and text and valid_evidence:
        _append_issue(
            issues,
            field,
            "generated",
            "\n".join(valid_evidence),
            f"{label}由AI根据简历已有信息生成，仅供参考，需人工确认",
        )
        return text, "generated", valid_evidence

    if allow_inferred and valid_evidence:
        return "\n".join(valid_evidence), "inferred", valid_evidence

    _append_issue(
        issues,
        field,
        "missing",
        "",
        f"简历中没有可核验或可安全生成的{label}内容",
    )
    return "", "missing", []


def _normalize_experience_list(
    raw_items: Any,
    field_prefix: str,
    source_markdown: str,
    issues: list[dict[str, str]],
    include_achievement: bool = False,
    allow_partial: bool = False,
) -> list[dict[str, str]]:
    """校验并标准化工作或项目经历。"""

    result = []
    is_project = field_prefix == "project_experiences"

    for index, raw_item in enumerate(_list(raw_items)):
        item = _mapping(raw_item)
        evidence = _mapping(item.get("evidence"))
        normalized: dict[str, str] = {}
        fields = (
            ("project_name", "项目名称"),
            ("position_name", "岗位名称"),
            ("description", "项目描述"),
        ) if is_project else (
            ("company_name", "公司名称"),
            ("position_name", "岗位名称"),
            ("description", "工作描述"),
        )

        if is_project and include_achievement:
            fields = (*fields, ("achievement", "工作业绩"))

        for field, label in fields:
            value = _text(item.get(field))
            field_evidence = _text(evidence.get(field))

            if value and _value_has_evidence(
                value,
                field_evidence,
                source_markdown,
            ):
                normalized[field] = value
            else:
                normalized[field] = ""

                if field != "position_name" or value:
                    _append_issue(
                        issues,
                        f"{field_prefix}[{index}].{field}",
                        "ambiguous" if field_evidence else "missing",
                        field_evidence,
                        f"没有可核验的{label}",
                    )

        original_time = _text(item.get("original_time"))
        time_evidence = _text(evidence.get("original_time"))

        if original_time and _value_has_evidence(
            original_time,
            time_evidence,
            source_markdown,
        ):
            normalized["original_time"] = original_time
            normalized.update(normalize_date_range(original_time))
        else:
            normalized["original_time"] = ""
            normalized.update(normalize_date_range(""))
            _append_issue(
                issues,
                f"{field_prefix}[{index}].time",
                "missing",
                time_evidence,
                "没有可核验的时间信息",
            )

        primary_field = (
            "project_name" if is_project else "company_name"
        )

        has_partial_content = any(
            normalized.get(field)
            for field in (
                "original_time",
                "company_name",
                "project_name",
                "position_name",
                "description",
                "achievement",
            )
        )

        if normalized.get(primary_field) or (
            allow_partial and has_partial_content
        ):
            result.append(normalized)

    return result


def _restore_explicit_project_blocks(
    projects: list[dict[str, str]],
    source_markdown: str,
    issues: list[dict[str, str]],
    include_achievement: bool = False,
) -> list[dict[str, str]]:
    """用Markdown中的完整项目块补回模型省略的项目正文。"""

    result = [dict(item) for item in projects]

    for project_name, project_body in _extract_explicit_project_blocks(
        source_markdown
    ):
        description_body = project_body
        recovered_achievement = ""
        if include_achievement:
            description_body, recovered_achievement = (
                _split_project_achievement(project_body)
            )
        description_body = _strip_project_description_label(
            description_body
        )
        matched_index = _find_project_index(result, project_name)

        if matched_index is None:
            recovered = {
                "project_name": project_name,
                "position_name": "",
                "original_time": "",
                **normalize_date_range(""),
                "description": description_body,
            }
            if include_achievement:
                recovered["achievement"] = recovered_achievement
            result.append(recovered)
            _append_issue(
                issues,
                f"project_experiences[{len(result) - 1}]",
                "recovered",
                project_name,
                "模型漏提该项目，已从恢复后Markdown完整补回",
            )
            continue

        existing_description = _text(
            result[matched_index].get("description")
        )
        complete_description = _merge_project_content(
            existing_description,
            description_body,
        )
        if complete_description != existing_description:
            result[matched_index]["description"] = complete_description
            _append_issue(
                issues,
                f"project_experiences[{matched_index}].description",
                "recovered",
                project_name,
                "模型未完整保留项目正文，已从恢复后Markdown补回",
            )

        if include_achievement and recovered_achievement:
            existing_achievement = _text(
                result[matched_index].get("achievement")
            )
            complete_achievement = _merge_project_content(
                existing_achievement,
                recovered_achievement,
            )
            if complete_achievement != existing_achievement:
                result[matched_index]["achievement"] = complete_achievement
                _append_issue(
                    issues,
                    f"project_experiences[{matched_index}].achievement",
                    "recovered",
                    project_name,
                    "模型未完整保留工作业绩，已从恢复后Markdown补回",
                )

    return result


def _extract_explicit_project_blocks(
    source_markdown: str,
) -> list[tuple[str, str]]:
    """提取项目分区中由下一级Markdown标题划分的完整项目块。"""

    lines = source_markdown.splitlines()
    headings: list[tuple[int, int, str]] = []
    heading_pattern = re.compile(r"^(#{1,6})[ \t]+(.+?)\s*$")

    for index, line in enumerate(lines):
        match = heading_pattern.match(line)
        if match:
            headings.append((index, len(match.group(1)), match.group(2)))

    blocks: list[tuple[str, str]] = []
    for heading_index, (line_index, level, title) in enumerate(headings):
        if not _is_project_section_heading(title):
            continue

        section_end = len(lines)
        for next_line, next_level, _ in headings[heading_index + 1 :]:
            if next_level <= level:
                section_end = next_line
                break

        project_headings = [
            (candidate_line, candidate_title)
            for candidate_line, candidate_level, candidate_title in headings
            if line_index < candidate_line < section_end
            and candidate_level == level + 1
            and not _is_project_content_subheading(candidate_title)
        ]

        for project_index, (project_line, project_title) in enumerate(
            project_headings
        ):
            block_end = (
                project_headings[project_index + 1][0]
                if project_index + 1 < len(project_headings)
                else section_end
            )
            body = "\n".join(lines[project_line + 1 : block_end]).strip()
            project_name = _project_name_from_heading(
                project_title,
                body,
            )
            if project_name:
                blocks.append((project_name, body))

    return blocks


def _is_project_section_heading(title: str) -> bool:
    """判断标题是否为项目经历分区。"""

    plain = _plain_markdown_heading(title)
    return bool(
        re.fullmatch(
            r"项目(?:经历|经验|案例)"
            r"(?:\s*[|｜/\-]?\s*Project\s+Experience)?",
            plain,
            flags=re.IGNORECASE,
        )
    )


def _is_project_content_subheading(title: str) -> bool:
    """避免把与项目标题同级的正文子标题误识别成新项目。"""

    plain = _plain_markdown_heading(title)
    return plain in {
        "项目描述",
        "应用技术",
        "主要职责",
        "工作内容",
        "功能说明",
        "实现细节",
        "工作业绩",
        "项目成果",
        "职责描述",
        "责任描述",
        "技术栈",
    }


def _project_name_from_heading(title: str, body: str) -> str:
    """从项目标题或正文中的项目名称标签取得名称。"""

    plain = _plain_markdown_heading(title)
    plain = re.sub(
        r"^项目(?:\d+|[一二三四五六七八九十百]+)\s*[:：、.\-]\s*",
        "",
        plain,
    )
    plain = re.split(r"[|｜]", plain, maxsplit=1)[0].strip()

    if re.fullmatch(r"项目(?:\d+|[一二三四五六七八九十百]+)", plain):
        name_match = re.search(
            r"^(?:\*\*)?项目名称\s*[:：](?:\*\*)?\s*(.+?)\s*$",
            body,
            flags=re.MULTILINE,
        )
        if name_match:
            plain = _plain_markdown_heading(name_match.group(1))

    return plain


def _plain_markdown_heading(value: str) -> str:
    """去除标题中的常见Markdown控制符。"""

    plain = value.strip()
    for pattern in (
        r"\*\*(.+?)\*\*",
        r"__(.+?)__",
        r"~~(.+?)~~",
        r"`([^`]+)`",
    ):
        plain = re.sub(pattern, r"\1", plain)
    return plain.strip().rstrip(":：")


def _find_project_index(
    projects: list[dict[str, str]],
    project_name: str,
) -> int | None:
    """按项目名称匹配模型结果与Markdown项目块。"""

    target = _comparable_project_name(project_name)
    if not target:
        return None

    for index, project in enumerate(projects):
        candidate = _comparable_project_name(
            _text(project.get("project_name"))
        )
        if not candidate:
            continue
        if candidate == target:
            return index
        if min(len(candidate), len(target)) >= 4 and (
            candidate in target or target in candidate
        ):
            return index

    return None


def _comparable_project_name(value: str) -> str:
    """生成仅用于项目名称匹配的保守比较文本。"""

    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", value.casefold())


def _merge_project_content(existing: str, complete_body: str) -> str:
    """优先采用完整项目块，并保留块外但可核验的模型描述。"""

    if not complete_body:
        return existing
    if not existing:
        return complete_body

    compact_existing = re.sub(r"\s+", "", existing)
    compact_body = re.sub(r"\s+", "", complete_body)
    if compact_existing in compact_body:
        return complete_body
    if compact_body in compact_existing:
        return existing
    return f"{complete_body}\n\n{existing}"


def _split_project_achievement(project_body: str) -> tuple[str, str]:
    """从奇瑞项目正文中分离明确标注的工作业绩，避免重复展示。"""

    lines = project_body.splitlines()
    achievement_start: int | None = None
    inline_achievement = ""
    label_pattern = re.compile(
        r"^\s*(?:#{1,6}\s*)?(?:\*\*)?工作业绩\s*[:：]"
        r"(?:\*\*)?\s*(.*?)\s*$"
    )

    for index, line in enumerate(lines):
        match = label_pattern.match(line)
        if match:
            achievement_start = index
            inline_achievement = match.group(1).strip()
            break

    if achievement_start is None:
        return project_body, ""

    achievement_end = len(lines)
    next_label_pattern = re.compile(
        r"^\s*(?:#{1,6}\s*)?(?:\*\*)?"
        r"(?:项目描述|应用技术|主要职责|功能说明|实现细节|项目成果)"
        r"\s*[:：](?:\*\*)?"
    )
    for index in range(achievement_start + 1, len(lines)):
        if next_label_pattern.match(lines[index]):
            achievement_end = index
            break

    achievement_lines = []
    if inline_achievement:
        achievement_lines.append(inline_achievement)
    achievement_lines.extend(lines[achievement_start + 1 : achievement_end])
    achievement = "\n".join(achievement_lines).strip()

    description_lines = (
        lines[:achievement_start] + lines[achievement_end:]
    )
    description = "\n".join(description_lines).strip()
    return description, achievement


def _strip_project_description_label(project_body: str) -> str:
    """移除项目块内层的描述标签，避免与Word固定标签重复。"""

    lines = project_body.splitlines()
    label_pattern = re.compile(
        r"^\s*(?:#{1,6}\s*)?(?:\*\*)?项目描述\s*[:：]"
        r"(?:\*\*)?\s*(.*?)\s*$"
    )
    for index, line in enumerate(lines):
        match = label_pattern.match(line)
        if not match:
            continue
        inline_description = match.group(1).strip()
        replacement = [inline_description] if inline_description else []
        lines[index : index + 1] = replacement
        break
    return "\n".join(lines).strip()


def _remove_resolved_project_issues(
    issues: list[dict[str, str]],
    projects: list[dict[str, str]],
) -> list[dict[str, str]]:
    """移除项目补回后已经失效的缺失或歧义记录。"""

    result = []
    pattern = re.compile(
        r"^project_experiences\[(\d+)\]\."
        r"(project_name|description|achievement)$"
    )
    for issue in issues:
        match = pattern.fullmatch(_text(issue.get("field")))
        if match and _text(issue.get("status")) in {"missing", "ambiguous"}:
            index = int(match.group(1))
            field = match.group(2)
            if index < len(projects) and _text(projects[index].get(field)):
                continue
        result.append(issue)
    return result


def _normalize_education_list(
    raw_items: Any,
    source_markdown: str,
    issues: list[dict[str, str]],
) -> list[dict[str, str]]:
    """校验教育经历，并应用教育年月业务规则。"""

    result = []

    for index, raw_item in enumerate(_list(raw_items)):
        item = _mapping(raw_item)
        evidence = _mapping(item.get("evidence"))
        normalized: dict[str, str] = {}

        for field, label in (
            ("school_name", "学校"),
            ("degree", "学历"),
            ("major", "专业"),
        ):
            value = _text(item.get(field))
            field_evidence = _text(evidence.get(field))

            if value and _value_has_evidence(
                value,
                field_evidence,
                source_markdown,
            ):
                if field == "degree":
                    value = _normalize_degree(value)

                normalized[field] = value

                if field == "degree" and not value:
                    _append_issue(
                        issues,
                        f"education_experiences[{index}].degree",
                        "ambiguous",
                        field_evidence,
                        "原文学历无法明确归入博士、硕士、本科或大专",
                    )
            else:
                normalized[field] = ""
                _append_issue(
                    issues,
                    f"education_experiences[{index}].{field}",
                    "ambiguous" if field_evidence else "missing",
                    field_evidence,
                    f"没有可核验的{label}信息",
                )

        original_time = _text(item.get("original_time"))
        time_evidence = _text(evidence.get("original_time"))

        if original_time and _value_has_evidence(
            original_time,
            time_evidence,
            source_markdown,
        ):
            normalized["original_time"] = original_time
            normalized.update(
                normalize_date_range(
                    original_time,
                    education=True,
                )
            )
        else:
            normalized["original_time"] = ""
            normalized.update(normalize_date_range("", education=True))
            _append_issue(
                issues,
                f"education_experiences[{index}].time",
                "missing",
                time_evidence,
                "没有可核验的教育时间",
            )

        if normalized.get("school_name"):
            result.append(normalized)

    return result


def _normalize_degree(value: str) -> str:
    """只把含义明确的学历映射到标准值。"""

    mappings = (
        ("博士", "博士"),
        ("硕士", "硕士"),
        ("本科", "本科"),
        ("学士", "本科"),
        ("大专", "大专"),
        ("专科", "大专"),
    )

    for keyword, normalized in mappings:
        if keyword in value:
            return normalized

    return ""


def _extract_date_tokens(text: str) -> list[str]:
    """按原文顺序提取年份或年月，不把年份区间误判为月份。"""

    tokens = []

    for match in re.finditer(
        r"(?<!\d)(\d{4})(?:\s*([./年-])\s*(\d{1,4})\s*月?)?",
        text,
    ):
        year = match.group(1)
        separator = match.group(2)
        trailing = match.group(3)

        if not trailing:
            tokens.append(year)
            continue

        number = int(trailing)

        if number <= 12:
            tokens.append(f"{year}/{number:02d}")
        elif separator == "-" and len(trailing) == 4:
            tokens.extend((year, trailing))
        else:
            tokens.append(year)

    return tokens[:2]


def _value_has_evidence(
    value: str,
    evidence: str,
    source_markdown: str,
) -> bool:
    """要求依据和字段值都能在原文中核验。"""

    return bool(
        evidence
        and _supported(evidence, source_markdown)
        and (
            _supported(value, evidence)
            or _supported(value, source_markdown)
        )
    )


def _supported(value: str, source: str) -> bool:
    """忽略空白和Markdown结构符号后检查原文包含关系。"""

    normalized_value = _comparable(value)
    return bool(
        normalized_value
        and normalized_value in _comparable(source)
    )


def _comparable(text: str) -> str:
    """生成仅用于证据匹配的保守规范化文本。"""

    text = re.sub(r"(?m)^\s*(?:#{1,6}|[-*+]|\d+[.、)])\s*", "", text)
    return re.sub(r"\s+", "", text)


def _date_sort_key(value: str) -> tuple[int, int]:
    """将日期转为排序键；至今视为最新。"""

    if value == "至今":
        return 9999, 12

    match = re.fullmatch(r"(\d{4})(?:/(\d{2}))?", value)

    if not match:
        return 0, 0

    return int(match.group(1)), int(match.group(2) or 0)


def _field_item(value: Any) -> dict[str, str]:
    """读取基本字段的值和证据。"""

    item = _mapping(value)
    return {
        "value": _text(item.get("value")),
        "evidence": _text(item.get("evidence")),
    }


def _append_issue(
    issues: list[dict[str, str]],
    field: str,
    status: str,
    evidence: str,
    note: str,
) -> None:
    """追加统一格式的缺失或歧义说明。"""

    issues.append(
        {
            "field": field,
            "status": status,
            "evidence": evidence,
            "note": note,
        }
    )


def _deduplicate_issues(
    issues: list[dict[str, str]],
) -> list[dict[str, str]]:
    """按全部字段去重并保持原顺序。"""

    seen = set()
    result = []

    for issue in issues:
        key = tuple(issue.values())

        if key not in seen:
            seen.add(key)
            result.append(issue)

    return result


def _mapping(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def _list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def _text(value: Any) -> str:
    return value.strip() if isinstance(value, str) else ""
