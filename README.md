# 文本对比工具 (text_compare_tools)

对比两段文本，把**不同的字符**在 Excel 中用**红色字体高亮**标出。
提供命令行工具和 Web 服务两种使用方式。

## 功能

- 字符级对比（基于 `difflib`），只标红真正不同的字符，相同部分保持黑色。
- 输出 Excel：两列分别放两段文本，不同字符红色高亮，并标注每行是否有差异。
- Web 端支持三种上传方式：
  1. **两个 txt 文件** —— 行一一对应，逐行对比。
  2. **Excel 文件** —— 指定要对比的两个列名（表头在第一行）。
  3. **jsonl 文件** —— 每行一个 json（如 `{"src":"...","mt":"..."}`），指定要对比的两个字段名。
- 结果提供下载链接。

## 安装

```bash
pip install -r requirements.txt
```

## Web 服务

```bash
python3 app.py
```

浏览器打开 http://127.0.0.1:5000 ，按页面提示上传文件、填写列名/字段名，即可下载高亮结果。

## 命令行（两个 txt 文件）

```bash
python3 text_diff.py file1.txt file2.txt -o diff_result.xlsx
```

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `app.py` | Flask 服务（三种上传方式 + 下载接口） |
| `diff_core.py` | 字符级 diff + 生成红色高亮 Excel 的核心逻辑 |
| `text_diff.py` | 命令行工具 |
| `templates/index.html` | 上传页面 |
