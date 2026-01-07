FROM python:3.11-slim

# 安装精简版LibreOffice（约500MB vs 完整版2GB+）
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer-nogui \
    libreoffice-calc-nogui \
    libreoffice-impress-nogui \
    fonts-wqy-zenhei \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean \
    && rm -rf /usr/share/libreoffice/help

# 设置LibreOffice环境变量
ENV HOME=/app
ENV LANG=C.UTF-8
ENV LC_ALL=C.UTF-8

# 设置工作目录
WORKDIR /app

# 复制依赖文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY . .

# 创建临时文件目录
RUN mkdir -p /app/temp_files

# 暴露端口
EXPOSE 5000

# 使用gunicorn生产服务器
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--threads", "4", "--timeout", "120", "app:app"]
