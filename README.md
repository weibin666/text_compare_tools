# 文本对比工具 (text_compare_tools)

逐行对比两段文本，把**不同的字符高亮**标出（<span style="color:#d11">红=删除/缺失</span>，<span style="color:#080">绿=新增</span>），
支持**网页在线预览**并导出 Excel。提供命令行工具和 Web 服务两种使用方式。

## 功能

- **字符级 / 词级**对比（基于 `difflib`），只标出真正不同的部分。
- **删除/新增分色**：左侧删除标红，右侧新增标绿。
- **对比选项**：忽略大小写 / 空白 / 标点；只导出有差异的行。
- **统计信息**：每行相似度、整体差异行数、平均相似度。
- **网页内高亮预览**（前 200 行），无需下载即可查看；完整结果导出 Excel。
- **三种输入**：
  1. 两个 txt 文件（行一一对应）。
  2. Excel / CSV 文件（指定要对比的两个列名）。
  3. jsonl 文件（每行一个 json，如 `{"src":"...","mt":"..."}`，指定两个字段名）。

## 安装

```bash
pip install -r requirements.txt
```

## 本地运行（开发）

```bash
python3 app.py            # http://127.0.0.1:5000
FLASK_DEBUG=1 python3 app.py   # 需要调试时
```

## 生产部署

用 gunicorn 启动（不要用开发服务器 / debug 模式）：

```bash
gunicorn -w 4 -b 0.0.0.0:8000 --timeout 120 app:app
```

可用环境变量调整：

| 变量 | 默认 | 说明 |
| --- | --- | --- |
| `MAX_UPLOAD_MB` | 50 | 单次上传大小上限（MB） |
| `RESULT_TTL` | 1800 | 结果文件保留秒数，过期自动清理 |
| `RATE_LIMIT` | 30 | 每 IP 每分钟最大请求数 |

> 结果文件存放在 `results/`，由后台线程按 `RESULT_TTL` 自动清理；多进程部署时建议把它放在共享/独立存储或挂载卷。

## 命令行（两个 txt 文件）

```bash
python3 text_diff.py file1.txt file2.txt -o diff_result.xlsx \
    --granularity char --ignore-case --ignore-ws --ignore-punct --only-diff
```

## 测试

```bash
python3 test_diff_core.py        # 或 python3 -m pytest
```

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `app.py` | Flask 服务（三种上传方式、对比选项、预览、下载） |
| `diff_core.py` | 字符/词级 diff + 生成高亮 Excel + 预览 HTML 的核心逻辑 |
| `utils.py` | 结果清理、限流、文件校验 |
| `text_diff.py` | 命令行工具 |
| `templates/index.html` | 上传 + 预览页面 |
| `test_diff_core.py` | 单元测试 |
