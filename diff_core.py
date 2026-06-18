#!/usr/bin/env python3
"""
文本对比核心逻辑：字符级 diff + 生成带红色高亮的 Excel。

被命令行工具 text_diff.py 和 Flask 服务 app.py 共用。
"""

import difflib

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font, PatternFill, Alignment

# 红色字体（用于不同的字符）
RED = InlineFont(color="FFFF0000")
# 普通黑色字体（用于相同的字符）
BLACK = InlineFont(color="FF000000")


def _build_rich(text, diff_ranges):
    """根据需要标红的区间列表，构建 CellRichText（区间内红色，其余黑色）。"""
    if not text:
        return CellRichText("")

    flags = [False] * len(text)
    for start, end in diff_ranges:
        for i in range(start, end):
            flags[i] = True

    blocks = []
    cur = text[0]
    cur_flag = flags[0]
    for ch, flag in zip(text[1:], flags[1:]):
        if flag == cur_flag:
            cur += ch
        else:
            blocks.append(TextBlock(RED if cur_flag else BLACK, cur))
            cur = ch
            cur_flag = flag
    blocks.append(TextBlock(RED if cur_flag else BLACK, cur))
    return CellRichText(*blocks)


def _diff_line(a, b):
    """字符级对比两行，返回 (a 需标红区间, b 需标红区间)。"""
    a_ranges, b_ranges = [], []
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if tag in ("replace", "delete"):
            a_ranges.append((i1, i2))
        if tag in ("replace", "insert"):
            b_ranges.append((j1, j2))
    return a_ranges, b_ranges


def build_diff_workbook(pairs, header_a="文本1", header_b="文本2"):
    """
    pairs: [(a, b), ...] 待对比的成对字符串。
    返回一个 openpyxl Workbook，两列分别放 a/b，不同字符红色高亮。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "diff"

    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="FFD9D9D9")
    for col, title in enumerate(["行号", header_a, header_b, "是否不同"], start=1):
        c = ws.cell(row=1, column=col, value=title)
        c.font = header_font
        c.fill = header_fill

    wrap = Alignment(wrap_text=True, vertical="top")
    diff_count = 0

    for idx, (a, b) in enumerate(pairs):
        a = "" if a is None else str(a)
        b = "" if b is None else str(b)
        row = idx + 2
        ws.cell(row=row, column=1, value=idx + 1)

        if a == b:
            ws.cell(row=row, column=2, value=a)
            ws.cell(row=row, column=3, value=b)
            ws.cell(row=row, column=4, value="")
        else:
            diff_count += 1
            a_ranges, b_ranges = _diff_line(a, b)
            ws.cell(row=row, column=2).value = _build_rich(a, a_ranges)
            ws.cell(row=row, column=3).value = _build_rich(b, b_ranges)
            ws.cell(row=row, column=4, value="✗")

        for col in (2, 3):
            ws.cell(row=row, column=col).alignment = wrap

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 8
    ws.freeze_panes = "A2"

    return wb, diff_count
