#!/usr/bin/env python3
"""
文本对比核心逻辑：可配置粒度的字符/词级 diff，输出结构化片段，
供 Excel 生成（红=删除/缺失，绿=新增）和网页预览复用。

被命令行工具 text_diff.py 和 Flask 服务 app.py 共用。
"""

import html
import re
from difflib import SequenceMatcher

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font, PatternFill, Alignment

# 字体颜色：删除/缺失=红，新增=绿，相同=黑
RED = InlineFont(color="FFFF0000")
GREEN = InlineFont(color="FF008000")
BLACK = InlineFont(color="FF000000")

# 网页预览的颜色
_HTML_COLOR = {"del": "#d11", "ins": "#080"}

_WORD_RE = re.compile(r"\w+|\s+|[^\w\s]", re.UNICODE)
_PUNCT_RE = re.compile(r"[^\w\s]", re.UNICODE)
_WS_RE = re.compile(r"\s+")


def _tokenize(text, granularity):
    """按粒度切分文本。char=逐字符，word=词/空白/标点为单位。"""
    if granularity == "word":
        return _WORD_RE.findall(text)
    return list(text)


def _norm_key(tok, ignore_case, ignore_ws, ignore_punct):
    """生成用于比对的归一化 key；返回空串表示该 token 在比对中被忽略。"""
    k = tok
    if ignore_case:
        k = k.lower()
    if ignore_punct:
        k = _PUNCT_RE.sub("", k)
    if ignore_ws:
        k = _WS_RE.sub("", k)
    return k


def _merge(tokens, ops):
    """把相邻、同类型的 token 合并成 (文本, 类型) 片段。"""
    if not tokens:
        return []
    segs = []
    cur_text = tokens[0]
    cur_op = ops[0]
    for tok, op in zip(tokens[1:], ops[1:]):
        if op == cur_op:
            cur_text += tok
        else:
            segs.append((cur_text, cur_op))
            cur_text, cur_op = tok, op
    segs.append((cur_text, cur_op))
    return segs


def diff_pair(a, b, granularity="char", ignore_case=False,
              ignore_ws=False, ignore_punct=False):
    """
    对比两段文本，返回 (a_segments, b_segments, similarity)。
    segments 为 [(文本, 类型), ...]，类型 ∈ {"equal","del","ins"}。
    similarity 为 0~1 的相似度。
    """
    a = "" if a is None else str(a)
    b = "" if b is None else str(b)
    ta, tb = _tokenize(a, granularity), _tokenize(b, granularity)
    ka = [_norm_key(t, ignore_case, ignore_ws, ignore_punct) for t in ta]
    kb = [_norm_key(t, ignore_case, ignore_ws, ignore_punct) for t in tb]

    # 过滤掉归一化后为空的 token（被忽略的空白/标点），它们一律按相同处理
    fa = [i for i, k in enumerate(ka) if k != ""]
    fb = [j for j, k in enumerate(kb) if k != ""]
    keys_a = [ka[i] for i in fa]
    keys_b = [kb[j] for j in fb]

    op_a = ["equal"] * len(ta)
    op_b = ["equal"] * len(tb)

    sm = SequenceMatcher(None, keys_a, keys_b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if tag in ("replace", "delete"):
            for x in range(i1, i2):
                op_a[fa[x]] = "del"
        if tag in ("replace", "insert"):
            for x in range(j1, j2):
                op_b[fb[x]] = "ins"

    if not keys_a and not keys_b:
        sim = 1.0
    else:
        sim = sm.ratio()

    return _merge(ta, op_a), _merge(tb, op_b), sim


def _rich(segments):
    """结构化片段 -> Excel 单元格富文本。"""
    if not segments:
        return CellRichText("")
    blocks = []
    for text, op in segments:
        font = RED if op == "del" else GREEN if op == "ins" else BLACK
        blocks.append(TextBlock(font, text))
    return CellRichText(*blocks)


def segments_to_html(segments):
    """结构化片段 -> 网页预览的 HTML（带颜色高亮，已转义）。"""
    parts = []
    for text, op in segments:
        esc = html.escape(text).replace("\n", "<br>")
        color = _HTML_COLOR.get(op)
        if color:
            parts.append(f'<span style="color:{color};font-weight:600">{esc}</span>')
        else:
            parts.append(esc)
    return "".join(parts)


def build_diff_workbook(pairs, header_a="文本1", header_b="文本2", *,
                        granularity="char", ignore_case=False, ignore_ws=False,
                        ignore_punct=False, only_diff=False, preview_limit=200):
    """
    pairs: [(a, b), ...] 待对比的成对字符串。
    返回 (workbook, result)，result 含统计信息与预览数据：
        {
          "total": 总行数, "diff_count": 差异行数,
          "avg_similarity": 平均相似度,
          "preview": [{"idx","sim","a_html","b_html","diff"}, ...]  # 前 preview_limit 行
        }
    """
    opts = dict(granularity=granularity, ignore_case=ignore_case,
                ignore_ws=ignore_ws, ignore_punct=ignore_punct)

    wb = Workbook()
    ws = wb.active
    ws.title = "diff"

    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="FFD9D9D9")
    for col, title in enumerate(["行号", header_a, header_b, "相似度", "是否不同"], start=1):
        c = ws.cell(row=1, column=col, value=title)
        c.font = header_font
        c.fill = header_fill

    wrap = Alignment(wrap_text=True, vertical="top")
    diff_count = 0
    sim_sum = 0.0
    preview = []
    out_row = 1  # excel 已写到的行（含表头）

    for idx, (a, b) in enumerate(pairs):
        seg_a, seg_b, sim = diff_pair(a, b, **opts)
        is_diff = any(op != "equal" for _, op in seg_a) or \
            any(op != "equal" for _, op in seg_b)
        sim_sum += sim
        if is_diff:
            diff_count += 1
        if only_diff and not is_diff:
            continue

        out_row += 1
        ws.cell(row=out_row, column=1, value=idx + 1)
        ws.cell(row=out_row, column=2).value = _rich(seg_a)
        ws.cell(row=out_row, column=3).value = _rich(seg_b)
        ws.cell(row=out_row, column=4, value=round(sim * 100, 1))
        ws.cell(row=out_row, column=5, value="✗" if is_diff else "")
        for col in (2, 3):
            ws.cell(row=out_row, column=col).alignment = wrap

        if len(preview) < preview_limit:
            preview.append({
                "idx": idx + 1,
                "sim": round(sim * 100, 1),
                "a_html": segments_to_html(seg_a),
                "b_html": segments_to_html(seg_b),
                "diff": is_diff,
            })

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 9
    ws.column_dimensions["E"].width = 8
    ws.freeze_panes = "A2"

    total = len(pairs)
    result = {
        "total": total,
        "diff_count": diff_count,
        "avg_similarity": round((sim_sum / total) * 100, 1) if total else 100.0,
        "preview": preview,
    }
    return wb, result
