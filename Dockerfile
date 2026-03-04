# 构建阶段
FROM m.daocloud.io/docker.io/library/python:3.11-slim AS builder

WORKDIR /app

# 配置pip使用国内镜像源并安装uv
RUN pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple/ && \
    pip config set global.trusted-host pypi.tuna.tsinghua.edu.cn && \
    pip install --no-cache-dir uv

# 复制依赖文件（优化缓存层）
COPY pyproject.toml uv.lock README.md ./

# 安装依赖
RUN uv pip install --system --no-cache -e . --index-url https://pypi.tuna.tsinghua.edu.cn/simple/

# 生产阶段
FROM m.daocloud.io/docker.io/library/python:3.11-slim

WORKDIR /app

# 从构建阶段复制已安装的包
COPY --from=builder /usr/local/lib/python3.11/site-packages /usr/local/lib/python3.11/site-packages
COPY --from=builder /usr/local/bin /usr/local/bin

# 复制应用代码、模板文件和配置文件
COPY src/ ./src/
COPY config/ ./config/

# 创建目录和用户，设置权限
RUN mkdir -p ./generated_files && \
    useradd --create-home --shell /bin/bash app && \
    chown -R app:app /app

USER app

ENV PYTHONPATH=/app/src \
    PYTHONUNBUFFERED=1

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/health')" || exit 1

CMD ["python", "-m", "uvicorn", "src.main:app", "--host", "0.0.0.0", "--port", "8000"]
