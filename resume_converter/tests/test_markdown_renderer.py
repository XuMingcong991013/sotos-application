"""
通用简历Markdown语义解析器的离线测试。
"""

from __future__ import annotations

import unittest

from resume_generation.markdown_renderer import (
    find_markdown_residues,
    parse_resume_markdown,
)


class MarkdownRendererTests(unittest.TestCase):
    """验证Markdown结构转换和技术文本防误伤。"""

    def test_parses_headings_bold_lists_and_nested_lists(self) -> None:
        """标题、粗体及多级列表转换为语义块，不保留标记字符。"""

        blocks = parse_resume_markdown(
            "#### 通用标题\n\n"
            "普通文字 **重点文字**。\n\n"
            "- 第一项\n"
            "  - 子项目\n"
            "- 第二项"
        )

        self.assertEqual(blocks[0].kind, "heading")
        self.assertEqual(blocks[0].level, 4)
        self.assertEqual(
            "".join(fragment.text for fragment in blocks[0].fragments),
            "通用标题",
        )
        bold_fragments = [
            fragment
            for block in blocks
            for fragment in block.fragments
            if fragment.bold
        ]
        self.assertEqual(
            [fragment.text for fragment in bold_fragments],
            ["重点文字"],
        )
        list_blocks = [
            block
            for block in blocks
            if block.kind == "bullet_list"
        ]
        self.assertEqual(len(list_blocks), 3)
        self.assertEqual([block.level for block in list_blocks], [0, 1, 0])

    def test_common_malformed_bold_boundaries_are_normalized(self) -> None:
        """多一个星号或闭合符前空格时仍转换为粗体。"""

        blocks = parse_resume_markdown(
            "***任意标签:**\n\n**另一标签： **"
        )
        fragments = [
            fragment
            for block in blocks
            for fragment in block.fragments
        ]

        self.assertEqual(
            [fragment.text for fragment in fragments],
            ["任意标签:", "另一标签："],
        )
        self.assertTrue(all(fragment.bold for fragment in fragments))

    def test_technical_symbols_are_not_treated_as_markdown(self) -> None:
        """不完整的语法边界不会破坏常见技术名称和代码符号。"""

        source = "C++、C#、char*、A*算法、#include、2**3"
        blocks = parse_resume_markdown(source)
        visible = "".join(
            fragment.text
            for block in blocks
            for fragment in block.fragments
        )

        self.assertEqual(visible, source)
        self.assertFalse(
            any(
                fragment.bold
                for block in blocks
                for fragment in block.fragments
            )
        )

    def test_links_code_quotes_strikethrough_and_tables_are_supported(self) -> None:
        """其他常见Markdown结构也不会把控制标记写入可见文字。"""

        blocks = parse_resume_markdown(
            "> 引用内容\n\n"
            "[链接文字](https://example.com) 和 `代码`、~~删除~~\n\n"
            "| 字段 | 内容 |\n"
            "| --- | --- |\n"
            "| A | B |"
        )
        visible = "\n".join(
            "".join(fragment.text for fragment in block.fragments)
            for block in blocks
        )

        self.assertIn("引用内容", visible)
        self.assertIn("链接文字", visible)
        self.assertIn("代码", visible)
        self.assertIn("删除", visible)
        self.assertIn("字段 | 内容", visible)
        self.assertTrue(
            any(
                fragment.link == "https://example.com"
                for block in blocks
                for fragment in block.fragments
            )
        )
        self.assertTrue(
            any(
                fragment.strike
                for block in blocks
                for fragment in block.fragments
            )
        )

    def test_residue_scan_uses_markdown_boundaries(self) -> None:
        """残留检查能发现结构标记但不会误报技术符号。"""

        residues = find_markdown_residues(
            [
                "#### 标题",
                "普通 **粗体**",
                "C++ C# char* A*算法 #include 2**3",
            ]
        )

        self.assertEqual(residues, ["#### 标题", "普通 **粗体**"])
