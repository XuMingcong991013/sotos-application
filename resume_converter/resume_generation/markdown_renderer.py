"""
把简历中的通用Markdown解析为可供Word生成器使用的语义块。

本模块只负责理解Markdown结构，不包含候选人、公司、项目或具体栏目词汇。
原始Markdown和结构化JSON保持不变，Word生成阶段根据这些语义块设置格式。
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass, replace
from typing import Any

from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode


@dataclass(frozen=True)
class InlineFragment:
    """一段具有相同内联格式的可见文字。"""

    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    code: bool = False
    link: str = ""


@dataclass(frozen=True)
class MarkdownBlock:
    """一个可转换为Word段落的Markdown块。"""

    kind: str
    fragments: tuple[InlineFragment, ...]
    level: int = 0
    list_id: int = 0
    list_start: int = 1
    quote_depth: int = 0


class _ResumeMarkdownParser:
    """将markdown-it语法树转换为简历Word所需的精简语义。"""

    def __init__(self) -> None:
        self._markdown = MarkdownIt(
            "commonmark",
            {"html": True},
        ).enable(["table", "strikethrough"])
        self._next_list_id = 1

    def parse(self, text: str) -> list[MarkdownBlock]:
        """解析常见块级和内联Markdown，忽略纯空白与图片。"""

        if not isinstance(text, str) or not text.strip():
            return []

        root = SyntaxTreeNode(
            self._markdown.parse(_normalize_common_markdown(text))
        )
        blocks: list[MarkdownBlock] = []
        self._walk_nodes(root.children, blocks, quote_depth=0)
        return [block for block in blocks if _block_has_content(block)]

    def _walk_nodes(
        self,
        nodes: list[SyntaxTreeNode],
        blocks: list[MarkdownBlock],
        quote_depth: int,
    ) -> None:
        for node in nodes:
            node_type = node.type

            if node_type == "paragraph":
                self._append_text_block(
                    node,
                    blocks,
                    "quote" if quote_depth else "paragraph",
                    quote_depth=quote_depth,
                )
            elif node_type == "heading":
                level = _heading_level(node)
                self._append_text_block(
                    node,
                    blocks,
                    "heading",
                    level=level,
                    quote_depth=quote_depth,
                )
            elif node_type in {"bullet_list", "ordered_list"}:
                self._walk_list(
                    node,
                    blocks,
                    ordered=node_type == "ordered_list",
                    level=0,
                    quote_depth=quote_depth,
                )
            elif node_type == "blockquote":
                self._walk_nodes(
                    node.children,
                    blocks,
                    quote_depth=quote_depth + 1,
                )
            elif node_type in {"fence", "code_block"}:
                content = _node_content(node).strip("\r\n")
                if content.strip():
                    blocks.append(
                        MarkdownBlock(
                            kind="code",
                            fragments=(
                                InlineFragment(content, code=True),
                            ),
                            quote_depth=quote_depth,
                        )
                    )
            elif node_type == "table":
                self._walk_table(node, blocks, quote_depth)
            elif node_type in {"html_block", "html_inline"}:
                visible = _visible_html(_node_content(node))
                if visible:
                    blocks.append(
                        MarkdownBlock(
                            kind="paragraph",
                            fragments=(InlineFragment(visible),),
                            quote_depth=quote_depth,
                        )
                    )
            elif node_type == "hr":
                # Markdown分隔线在标准简历模板中没有额外语义。
                continue
            elif node.children:
                self._walk_nodes(node.children, blocks, quote_depth)

    def _walk_list(
        self,
        node: SyntaxTreeNode,
        blocks: list[MarkdownBlock],
        ordered: bool,
        level: int,
        quote_depth: int,
    ) -> None:
        list_id = self._next_list_id
        self._next_list_id += 1
        list_start = int(node.attrs.get("start", 1) or 1)

        for item in node.children:
            first_text_block = True
            for child in item.children:
                if child.type in {"paragraph", "heading"}:
                    fragments = _inline_fragments(child)
                    if not _fragments_have_content(fragments):
                        continue
                    kind = (
                        "ordered_list"
                        if ordered and first_text_block
                        else "bullet_list"
                        if first_text_block
                        else "paragraph"
                    )
                    blocks.append(
                        MarkdownBlock(
                            kind=kind,
                            fragments=tuple(fragments),
                            level=min(level, 8),
                            list_id=list_id,
                            list_start=list_start,
                            quote_depth=quote_depth,
                        )
                    )
                    first_text_block = False
                elif child.type in {"bullet_list", "ordered_list"}:
                    self._walk_list(
                        child,
                        blocks,
                        ordered=child.type == "ordered_list",
                        level=level + 1,
                        quote_depth=quote_depth,
                    )
                else:
                    self._walk_nodes(
                        [child],
                        blocks,
                        quote_depth=quote_depth,
                    )

    def _walk_table(
        self,
        node: SyntaxTreeNode,
        blocks: list[MarkdownBlock],
        quote_depth: int,
    ) -> None:
        rows = _descendants_of_type(node, "tr")

        for row_index, row in enumerate(rows):
            cells = [
                child
                for child in row.children
                if child.type in {"th", "td"}
            ]
            fragments: list[InlineFragment] = []

            for cell_index, cell in enumerate(cells):
                if cell_index:
                    fragments.append(InlineFragment(" | "))
                cell_fragments = _inline_fragments(cell)
                if row_index == 0:
                    cell_fragments = [
                        replace(fragment, bold=True)
                        for fragment in cell_fragments
                    ]
                fragments.extend(cell_fragments)

            if _fragments_have_content(fragments):
                blocks.append(
                    MarkdownBlock(
                        kind="table_row",
                        fragments=tuple(_merge_fragments(fragments)),
                        quote_depth=quote_depth,
                    )
                )

    @staticmethod
    def _append_text_block(
        node: SyntaxTreeNode,
        blocks: list[MarkdownBlock],
        kind: str,
        level: int = 0,
        quote_depth: int = 0,
    ) -> None:
        fragments = _inline_fragments(node)
        if _fragments_have_content(fragments):
            blocks.append(
                MarkdownBlock(
                    kind=kind,
                    fragments=tuple(fragments),
                    level=level,
                    quote_depth=quote_depth,
                )
            )


def parse_resume_markdown(text: str) -> list[MarkdownBlock]:
    """公开的通用简历Markdown解析函数。"""

    return _ResumeMarkdownParser().parse(text)


def find_markdown_residues(texts: list[str]) -> list[str]:
    """查找仍可能作为普通文字泄漏到Word中的常见Markdown结构。"""

    patterns = (
        re.compile(r"^\s*#{1,6}\s+"),
        re.compile(r"\*\*\S(?:.*?\S)?\*\*"),
        re.compile(r"__\S(?:.*?\S)?__"),
        re.compile(r"~~\S(?:.*?\S)?~~"),
        re.compile(r"`[^`\r\n]+`"),
        re.compile(r"!?\[[^\]]*\]\([^)]*\)"),
        re.compile(r"^\s*>\s+"),
        re.compile(r"^\s*[-+*]\s+"),
    )
    residues: list[str] = []

    for text in texts:
        for line in text.splitlines() or [text]:
            if any(pattern.search(line) for pattern in patterns):
                residues.append(line)
                break

    return residues


def _inline_fragments(node: SyntaxTreeNode) -> list[InlineFragment]:
    inline_nodes = _descendants_of_type(node, "inline")
    if inline_nodes:
        source_nodes = inline_nodes[0].children
    else:
        source_nodes = node.children

    fragments: list[InlineFragment] = []
    _walk_inline(source_nodes, fragments, {})
    return _merge_fragments(fragments)


def _walk_inline(
    nodes: list[SyntaxTreeNode],
    fragments: list[InlineFragment],
    style: dict[str, Any],
) -> None:
    for node in nodes:
        node_type = node.type

        if node_type == "text":
            fragments.append(InlineFragment(_node_content(node), **style))
        elif node_type in {"softbreak", "hardbreak"}:
            fragments.append(InlineFragment("\n", **style))
        elif node_type == "code_inline":
            fragments.append(
                InlineFragment(_node_content(node), code=True, **style)
            )
        elif node_type == "strong":
            _walk_inline(node.children, fragments, {**style, "bold": True})
        elif node_type == "em":
            _walk_inline(node.children, fragments, {**style, "italic": True})
        elif node_type == "s":
            _walk_inline(node.children, fragments, {**style, "strike": True})
        elif node_type == "link":
            href = str(node.attrs.get("href", "") or "")
            _walk_inline(node.children, fragments, {**style, "link": href})
        elif node_type == "image":
            # 最终标准简历不从Markdown下载或插入图片。
            continue
        elif node_type in {"html_inline", "html_block"}:
            visible = _visible_html(_node_content(node))
            if visible:
                fragments.append(InlineFragment(visible, **style))
        elif node.children:
            _walk_inline(node.children, fragments, style)
        else:
            content = _node_content(node)
            if content:
                fragments.append(InlineFragment(content, **style))


def _merge_fragments(
    fragments: list[InlineFragment],
) -> list[InlineFragment]:
    merged: list[InlineFragment] = []

    for fragment in fragments:
        if not fragment.text:
            continue
        if merged and replace(merged[-1], text="") == replace(fragment, text=""):
            previous = merged[-1]
            merged[-1] = replace(previous, text=previous.text + fragment.text)
        else:
            merged.append(fragment)

    return merged


def _fragments_have_content(fragments: list[InlineFragment]) -> bool:
    return any(fragment.text.strip() for fragment in fragments)


def _block_has_content(block: MarkdownBlock) -> bool:
    return _fragments_have_content(list(block.fragments))


def _descendants_of_type(
    node: SyntaxTreeNode,
    node_type: str,
) -> list[SyntaxTreeNode]:
    result: list[SyntaxTreeNode] = []

    for child in node.children:
        if child.type == node_type:
            result.append(child)
        result.extend(_descendants_of_type(child, node_type))

    return result


def _heading_level(node: SyntaxTreeNode) -> int:
    tag = str(getattr(node, "tag", "") or "")
    if len(tag) == 2 and tag.startswith("h") and tag[1].isdigit():
        return int(tag[1])
    return 1


def _node_content(node: SyntaxTreeNode) -> str:
    return str(getattr(node, "content", "") or "")


def _visible_html(value: str) -> str:
    without_tags = re.sub(r"<[^>]+>", "", value)
    return html.unescape(without_tags).strip()


def _normalize_common_markdown(value: str) -> str:
    """修正常见的非标准粗体边界，不接触普通技术符号。"""

    def normalize_asterisks(match: re.Match[str]) -> str:
        content = match.group(2).strip()
        return f"**{content}**"

    def normalize_underscores(match: re.Match[str]) -> str:
        content = match.group(2).strip()
        return f"__{content}__"

    value = re.sub(
        r"(?<!\*)(\*{2,})([^*\r\n]+?)(\*{2,})(?!\*)",
        normalize_asterisks,
        value,
    )
    return re.sub(
        r"(?<!_)(_{2,})([^_\r\n]+?)(_{2,})(?!_)",
        normalize_underscores,
        value,
    )
