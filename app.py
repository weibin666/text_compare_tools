#!/usr/bin/env python3
"""
文本对比 Flask 服务。

支持三种上传方式：
  1. 两个 txt 文件（行一一对应）
  2. 一个 Excel / CSV 文件（指定两个列名）
  3. 一个 jsonl 文件（每行一个 json，指定两个字段名）

可配置对比粒度（字符/词）、忽略大小写/空白/标点、只导出差异行；
返回网页内高亮预览 + 统计信息 + 高亮 Excel 的下载链接。
"""

import csv
import io
import json
import logging
import os
import uuid

from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import load_workbook

from diff_core import build_diff_workbook
from utils import start_cleanup_thread, RateLimiter, ext_ok

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "50")) * 1024 * 1024

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger("text_diff")

# 生成的结果文件临时存放目录
RESULT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results")
os.makedirs(RESULT_DIR, exist_ok=True)

RESULT_TTL = int(os.environ.get("RESULT_TTL", "1800"))  # 结果文件保留秒数
start_cleanup_thread(RESULT_DIR, ttl_seconds=RESULT_TTL)
rate_limiter = RateLimiter(max_calls=int(os.environ.get("RATE_LIMIT", "30")), window=60)

ALLOWED_EXT = {
    "txt": {"txt", "text"},
    "excel": {"xlsx", "xlsm", "csv"},
    "jsonl": {"json", "jsonl", "txt"},
}


def _get_opts(form):
    """从表单读取对比选项。"""
    granularity = form.get("granularity", "char")
    if granularity not in ("char", "word"):
        granularity = "char"
    return dict(
        granularity=granularity,
        ignore_case=form.get("ignore_case") == "1",
        ignore_ws=form.get("ignore_ws") == "1",
        ignore_punct=form.get("ignore_punct") == "1",
        only_diff=form.get("only_diff") == "1",
    )


def _read_txt_lines(file_storage):
    text = file_storage.read().decode("utf-8-sig")
    return [line.rstrip("\r\n") for line in text.splitlines()]


def parse_txt(form, files):
    f1, f2 = files.get("txt1"), files.get("txt2")
    if not f1 or not f2 or not f1.filename or not f2.filename:
        raise ValueError("请同时上传两个 txt 文件。")
    for f in (f1, f2):
        if not ext_ok(f.filename, ALLOWED_EXT["txt"]):
            raise ValueError(f"不支持的文件类型：{f.filename}")
    lines1, lines2 = _read_txt_lines(f1), _read_txt_lines(f2)
    n = max(len(lines1), len(lines2))
    pairs = [(lines1[i] if i < len(lines1) else "",
              lines2[i] if i < len(lines2) else "") for i in range(n)]
    return pairs, f1.filename, f2.filename


def _read_table_rows(file_storage):
    """读取 Excel 或 CSV，返回 (表头list, 数据行迭代器)。"""
    name = file_storage.filename.lower()
    data = file_storage.read()
    if name.endswith(".csv"):
        text = data.decode("utf-8-sig")
        reader = csv.reader(io.StringIO(text))
        rows = list(reader)
        if not rows:
            raise ValueError("文件为空。")
        return [str(h).strip() for h in rows[0]], rows[1:]
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    try:
        header = list(next(it))
    except StopIteration:
        raise ValueError("文件为空。")
    header = ["" if h is None else str(h).strip() for h in header]
    return header, it


def parse_excel(form, files):
    f = files.get("excel")
    if not f or not f.filename:
        raise ValueError("请上传一个 Excel / CSV 文件。")
    if not ext_ok(f.filename, ALLOWED_EXT["excel"]):
        raise ValueError(f"不支持的文件类型：{f.filename}")
    col_a = (form.get("excel_col_a") or "").strip()
    col_b = (form.get("excel_col_b") or "").strip()
    if not col_a or not col_b:
        raise ValueError("请填写要对比的两个列名。")

    header, rows = _read_table_rows(f)
    if col_a not in header or col_b not in header:
        raise ValueError(f"列名不存在。文件中的列为: {', '.join(h for h in header if h)}")
    ia, ib = header.index(col_a), header.index(col_b)

    pairs = []
    for row in rows:
        row = list(row)
        a = row[ia] if ia < len(row) else ""
        b = row[ib] if ib < len(row) else ""
        pairs.append((a, b))
    return pairs, col_a, col_b


def parse_jsonl(form, files):
    f = files.get("jsonl")
    if not f or not f.filename:
        raise ValueError("请上传一个 jsonl 文件。")
    if not ext_ok(f.filename, ALLOWED_EXT["jsonl"]):
        raise ValueError(f"不支持的文件类型：{f.filename}")
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
    if not rate_limiter.allow(request.remote_addr or "?"):
        return jsonify(ok=False, error="请求过于频繁，请稍后再试。"), 429

    mode = request.form.get("mode")
    parser = PARSERS.get(mode)
    if not parser:
        return jsonify(ok=False, error="未知的对比类型。"), 400
    try:
        pairs, header_a, header_b = parser(request.form, request.files)
        opts = _get_opts(request.form)
        wb, result = build_diff_workbook(pairs, header_a, header_b, **opts)

        token = uuid.uuid4().hex
        wb.save(os.path.join(RESULT_DIR, f"{token}.xlsx"))
        logger.info("diff mode=%s total=%s diff=%s", mode, result["total"], result["diff_count"])

        return jsonify(
            ok=True,
            total=result["total"],
            diff_count=result["diff_count"],
            avg_similarity=result["avg_similarity"],
            header_a=header_a,
            header_b=header_b,
            preview=result["preview"],
            download_url=f"/download/{token}",
        )
    except ValueError as e:
        return jsonify(ok=False, error=str(e)), 400
    except Exception as e:  # noqa: BLE001
        logger.exception("diff failed")
        return jsonify(ok=False, error=f"处理失败：{e}"), 500


@app.route("/download/<token>")
def download(token):
    if not token or not all(c in "0123456789abcdef" for c in token):
        return "非法请求", 400
    path = os.path.join(RESULT_DIR, f"{token}.xlsx")
    if not os.path.exists(path):
        return "文件不存在或已过期", 404
    return send_file(path, as_attachment=True, download_name="diff_result.xlsx")


if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="127.0.0.1", port=5000, debug=debug)
