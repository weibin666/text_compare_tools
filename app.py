#!/usr/bin/env python3
"""
文本对比 Flask 服务。

支持三种上传方式：
  1. 两个 txt 文件（行一一对应）
  2. 一个 Excel 文件（指定两个列名）
  3. 一个 jsonl 文件（每行一个 json，指定两个字段名，如 src / mt）

输出带红色高亮的 Excel，提供下载链接。
"""

import io
import json
import os
import uuid

from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import load_workbook

from diff_core import build_diff_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# 生成的结果文件临时存放目录
RESULT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results")
os.makedirs(RESULT_DIR, exist_ok=True)


def _read_txt(file_storage):
    """读取一个 txt 文件，返回去掉行尾换行的行列表。"""
    text = file_storage.read().decode("utf-8-sig")
    return [line.rstrip("\r\n") for line in text.splitlines()]


def parse_txt(form, files):
    f1 = files.get("txt1")
    f2 = files.get("txt2")
    if not f1 or not f2 or not f1.filename or not f2.filename:
        raise ValueError("请同时上传两个 txt 文件。")
    lines1 = _read_txt(f1)
    lines2 = _read_txt(f2)
    n = max(len(lines1), len(lines2))
    pairs = [(lines1[i] if i < len(lines1) else "",
              lines2[i] if i < len(lines2) else "") for i in range(n)]
    return pairs, f1.filename, f2.filename


def parse_excel(form, files):
    f = files.get("excel")
    if not f or not f.filename:
        raise ValueError("请上传一个 Excel 文件。")
    col_a = (form.get("excel_col_a") or "").strip()
    col_b = (form.get("excel_col_b") or "").strip()
    if not col_a or not col_b:
        raise ValueError("请填写要对比的两个列名。")

    wb = load_workbook(io.BytesIO(f.read()), data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        header = list(next(rows))
    except StopIteration:
        raise ValueError("Excel 文件为空。")
    header = ["" if h is None else str(h).strip() for h in header]

    if col_a not in header or col_b not in header:
        raise ValueError(f"列名不存在。文件中的列为: {', '.join(h for h in header if h)}")
    ia, ib = header.index(col_a), header.index(col_b)

    pairs = []
    for row in rows:
        a = row[ia] if ia < len(row) else ""
        b = row[ib] if ib < len(row) else ""
        pairs.append((a, b))
    return pairs, col_a, col_b


def parse_jsonl(form, files):
    f = files.get("jsonl")
    if not f or not f.filename:
        raise ValueError("请上传一个 jsonl 文件。")
    key_a = (form.get("json_key_a") or "").strip()
    key_b = (form.get("json_key_b") or "").strip()
    if not key_a or not key_b:
        raise ValueError("请填写要对比的两个字段名（如 src / mt）。")

    text = f.read().decode("utf-8-sig")
    pairs = []
    for lineno, line in enumerate(text.splitlines(), start=1):
        line = line.strip()
        if not line:
            continue
        try:
            obj = json.loads(line)
        except json.JSONDecodeError as e:
            raise ValueError(f"第 {lineno} 行不是合法 json：{e}")
        pairs.append((obj.get(key_a, ""), obj.get(key_b, "")))
    if not pairs:
        raise ValueError("文件中没有有效的 json 行。")
    return pairs, key_a, key_b


PARSERS = {"txt": parse_txt, "excel": parse_excel, "jsonl": parse_jsonl}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/diff", methods=["POST"])
def diff():
    mode = request.form.get("mode")
    parser = PARSERS.get(mode)
    if not parser:
        return jsonify(ok=False, error="未知的对比类型。"), 400
    try:
        pairs, header_a, header_b = parser(request.form, request.files)
        wb, diff_count = build_diff_workbook(pairs, header_a, header_b)

        token = uuid.uuid4().hex
        out_path = os.path.join(RESULT_DIR, f"{token}.xlsx")
        wb.save(out_path)

        return jsonify(
            ok=True,
            total=len(pairs),
            diff_count=diff_count,
            download_url=f"/download/{token}",
        )
    except ValueError as e:
        return jsonify(ok=False, error=str(e)), 400
    except Exception as e:  # noqa: BLE001
        return jsonify(ok=False, error=f"处理失败：{e}"), 500


@app.route("/download/<token>")
def download(token):
    # token 仅允许十六进制，防止路径穿越
    if not all(c in "0123456789abcdef" for c in token):
        return "非法请求", 400
    path = os.path.join(RESULT_DIR, f"{token}.xlsx")
    if not os.path.exists(path):
        return "文件不存在或已过期", 404
    return send_file(path, as_attachment=True, download_name="diff_result.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
