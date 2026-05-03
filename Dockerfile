# ─── 基础镜像 ─────────────────────────────────────
FROM python:3.12-slim

LABEL maintainer="xx调剂定制班"
LABEL description="考研调剂规划自动化生成工具 - Web 版"

# 设置环境变量
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TZ=Asia/Shanghai

# 安装系统依赖（openpyxl 和 pandas 可能需要）
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        curl \
    && rm -rf /var/lib/apt/lists/*

# 设置工作目录
WORKDIR /app

# ─── 安装 Python 依赖 ─────────────────────────────
COPY web/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# ─── 复制项目文件 ──────────────────────────────────
# 核心业务模块
COPY read.py ai_tool.py ./
# 默认模板文件（用户未上传模板时使用）
COPY 模板.xlsx ./
# Web 应用
COPY web/app.py ./app.py
COPY web/templates/ ./templates/

# ─── 创建数据目录（用于 Volume 挂载） ──────────────
RUN mkdir -p /app/uploads /app/output

# ─── 健康检查 ─────────────────────────────────────
HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5000/ || exit 1

# ─── 暴露端口 & 启动 ─────────────────────────────
EXPOSE 5000

CMD ["python", "app.py"]
