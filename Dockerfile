FROM python:3.12-slim

WORKDIR /app

# 先装依赖，利用构建缓存
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 结果文件目录
RUN mkdir -p results

EXPOSE 8000

# 生产用 gunicorn；worker / 超时可按需调整
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:8000", "--timeout", "120", "app:app"]
