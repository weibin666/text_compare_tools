# 文本对比工具 (text_compare_tools)

逐行对比两段文本，把**不同的字符高亮**标出（<span style="color:#d11">红=删除/缺失</span>，<span style="color:#080">绿=新增</span>），
支持**网页在线预览**并导出 Excel。提供命令行工具和 Web 服务两种使用方式。

## 功能

- **字符级 / 词级**对比（基于 `difflib`），只标出真正不同的部分。
- **删除/新增分色**：左侧删除标红，右侧新增标绿。
- **对比选项**：忽略大小写 / 空白 / 标点；只导出有差异的行。
- **统计信息**：每行相似度、整体差异行数、平均相似度。
- **网页内高亮预览**（前 200 行），无需下载即可查看。
- **多种导出**：高亮 Excel（多列对比时一对一个 sheet）+ HTML 报告（浏览器内「打印 / 存为 PDF」即可导出 PDF）。
- **四种输入**：
  1. 两个 txt 文件（行一一对应）。
  2. 单个 Excel / CSV 文件：可指定**一对或多对列名**，一次完成多列对比。
  3. 两个 Excel / CSV 文件：各取一列，按行对比。
  4. jsonl 文件（每行一个 json，如 `{"src":"...","mt":"..."}`），可指定一对或多对字段名。

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

### 方式一：gunicorn（不要用开发服务器 / debug 模式）

```bash
gunicorn -w 4 -b 0.0.0.0:8000 --timeout 120 app:app
```

### 方式二：Docker + Nginx（HTTPS）

```bash
# 1) 把证书放到 nginx/certs/（fullchain.pem、privkey.pem）
#    没有正式证书时可先用自签名证书测试：
#    openssl req -x509 -newkey rsa:2048 -nodes -days 365 \
#      -keyout nginx/certs/privkey.pem -out nginx/certs/fullchain.pem -subj "/CN=localhost"
# 2) 启动
docker compose up -d --build
```

`app` 容器用 gunicorn 跑在内网，由 `nginx` 容器做反向代理并终止 HTTPS（80 自动跳 443）。
正式环境建议用 Let's Encrypt（certbot）签发证书。

### 环境变量

| 变量 | 默认 | 说明 |
| --- | --- | --- |
| `MAX_UPLOAD_MB` | 50 | 单次上传大小上限（MB） |
| `RESULT_TTL` | 1800 | 结果文件保留秒数，过期自动清理 |
| `RATE_LIMIT` | 30 | 每 IP 每分钟最大请求数 |
| `TRUST_PROXY` | 0 | 置 1 时从 `X-Forwarded-For` 取真实客户端 IP（反代后需开启） |
| `FLASK_DEBUG` | 0 | 仅本地调试用，生产保持 0 |

> 结果文件存放在 `results/`，由后台线程按 `RESULT_TTL` 自动清理；docker-compose 已用具名卷 `results` 持久化。多进程/多实例部署时建议放共享存储。
>
> **关于 PDF**：报告以打印优化的 HTML 形式提供，浏览器「打印 → 存为 PDF」即可得到 PDF，无需服务端重型依赖。若需服务端直接生成 PDF，可另行引入 WeasyPrint。

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
| `app.py` | Flask 服务（四种上传方式、对比选项、预览、Excel/报告下载） |
| `diff_core.py` | 字符/词级 diff + 多 sheet Excel + 预览/HTML 报告的核心逻辑 |
| `utils.py` | 结果清理、限流、文件校验 |
| `text_diff.py` | 命令行工具 |
| `templates/index.html` | 上传 + 预览页面 |
| `test_diff_core.py` | 单元测试 |
| `Dockerfile` / `docker-compose.yml` / `nginx/` | 容器化部署 + Nginx HTTPS 反代 |
