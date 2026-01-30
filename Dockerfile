# 稳定版Dockerfile - 解决SSL和网络问题
# 稳定版Dockerfile - 解决SSL和网络问题
FROM m.daocloud.io/docker.io/library/python:3.11-slim AS builder

# 设置工作目录
WORKDIR /app

# 设置pip配置，使用国内镜像源
RUN pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple/ && \
    pip config set global.trusted-host pypi.tuna.tsinghua.edu.cn

# 安装uv (使用国内镜像源，禁用SSL验证)
RUN pip install --no-cache-dir --trusted-host pypi.tuna.tsinghua.edu.cn uv

# 复制项目文件
COPY pyproject.toml uv.lock README.md ./

# 使用uv安装依赖 (配置国内镜像源)
RUN uv pip install --system --no-cache -e . --index-url https://pypi.tuna.tsinghua.edu.cn/simple/

# 生产阶段
FROM m.daocloud.io/docker.io/library/python:3.11-slim

# 设置工作目录
WORKDIR /app

# 从构建阶段复制已安装的包
COPY --from=builder /usr/local/lib/python3.11/site-packages /usr/local/lib/python3.11/site-packages
COPY --from=builder /usr/local/bin /usr/local/bin

# 复制应用代码、模板文件和配置文件
COPY src/ ./src/
COPY config/ ./config/

# 创建必要的目录
RUN mkdir -p ./generated_files

# 创建非root用户并设置权限
RUN useradd --create-home --shell /bin/bash app \
    && chown -R app:app /app

# 切换到非root用户
USER app

# 设置环境变量
ENV PYTHONPATH=/app/src
ENV PYTHONUNBUFFERED=1

# 暴露端口
EXPOSE 8000

# 健康检查 (使用Python内置模块)
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/health')" || exit 1

# 启动命令
CMD ["python", "-m", "uvicorn", "src.main:app", "--host", "0.0.0.0", "--port", "8000"]
