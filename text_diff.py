#!/usr/bin/env python3
"""
比较两个行数对应的 txt 文本，把每对行中不同的字符找出来，
输出到 Excel：两列分别放两个文件的内容，不同的字符用红色字体高亮。

用法:
    python3 text_diff.py file1.txt file2.txt [-o output.xlsx]
"""

import argparse
import sys

from diff_core import build_diff_workbook


def read_lines(path):
    """读取文件所有行，去掉行尾换行符。"""
    with open(path, "r", encoding="utf-8-sig") as f:
        return [line.rstrip("\r\n") for line in f]


def main():
    parser = argparse.ArgumentParser(description="比较两个 txt 文本并输出带红色高亮的 Excel")
    parser.add_argument("file1", help="第一个 txt 文件")
    parser.add_argument("file2", help="第二个 txt 文件")
    parser.add_argument("-o", "--output", default="diff_result.xlsx", help="输出的 xlsx 文件名")
    args = parser.parse_args()

    lines1 = read_lines(args.file1)
    lines2 = read_lines(args.file2)
    n1, n2 = len(lines1), len(lines2)
    if n1 != n2:
        print(f"警告: 两个文件行数不同 (file1={n1}, file2={n2})，按较多的行数对齐，缺失行视为空。",
              file=sys.stderr)
    n = max(n1, n2)

    pairs = [(lines1[i] if i < n1 else "", lines2[i] if i < n2 else "") for i in range(n)]
    wb, diff_count = build_diff_workbook(pairs, args.file1, args.file2)
    wb.save(args.output)
    print(f"完成：共 {n} 行，其中 {diff_count} 行有差异，结果已写入 {args.output}")


if __name__ == "__main__":
    main()
