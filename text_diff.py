#!/usr/bin/env python3
"""
比较两个行数对应的 txt 文本，把每对行中不同的字符找出来，
输出到 Excel：两列分别放两个文件的内容，删除/缺失红色、新增绿色高亮。

用法:
    python3 text_diff.py file1.txt file2.txt [-o output.xlsx]
        [--granularity char|word] [--ignore-case] [--ignore-ws]
        [--ignore-punct] [--only-diff]
"""

import argparse
import sys

from diff_core import build_diff_workbook


def read_lines(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        return [line.rstrip("\r\n") for line in f]


def main():
    p = argparse.ArgumentParser(description="比较两个 txt 文本并输出带高亮的 Excel")
    p.add_argument("file1", help="第一个 txt 文件")
    p.add_argument("file2", help="第二个 txt 文件")
    p.add_argument("-o", "--output", default="diff_result.xlsx", help="输出的 xlsx 文件名")
    p.add_argument("--granularity", choices=["char", "word"], default="char", help="对比粒度")
    p.add_argument("--ignore-case", action="store_true", help="忽略大小写")
    p.add_argument("--ignore-ws", action="store_true", help="忽略空白")
    p.add_argument("--ignore-punct", action="store_true", help="忽略标点")
    p.add_argument("--only-diff", action="store_true", help="只导出有差异的行")
    args = p.parse_args()

    lines1, lines2 = read_lines(args.file1), read_lines(args.file2)
    n1, n2 = len(lines1), len(lines2)
    if n1 != n2:
        print(f"警告: 两个文件行数不同 (file1={n1}, file2={n2})，按较多的行数对齐，缺失行视为空。",
              file=sys.stderr)
    n = max(n1, n2)
    pairs = [(lines1[i] if i < n1 else "", lines2[i] if i < n2 else "") for i in range(n)]

    wb, result = build_diff_workbook(
        pairs, args.file1, args.file2,
        granularity=args.granularity, ignore_case=args.ignore_case,
        ignore_ws=args.ignore_ws, ignore_punct=args.ignore_punct,
        only_diff=args.only_diff,
    )
    wb.save(args.output)
    print(f"完成：共 {result['total']} 行，差异 {result['diff_count']} 行，"
          f"平均相似度 {result['avg_similarity']}%，结果已写入 {args.output}")


if __name__ == "__main__":
    main()
