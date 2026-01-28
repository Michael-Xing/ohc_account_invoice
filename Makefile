# OHC账票生成FastAPI服务 Makefile

.PHONY: help install dev test clean docker-build docker-run docker-stop lint format docker-buildx k8s-deploy k8s-undeploy k8s-status k8s-logs k8s-shell

# 项目信息
PROJECT_NAME = ohc-account-invoice
VERSION = 1.0.0
REGISTRY = release.daocloud.io/dak
IMAGE_NAME = $(if $(REGISTRY),$(REGISTRY)/$(PROJECT_NAME),$(PROJECT_NAME))
FULL_IMAGE_NAME = $(IMAGE_NAME):$(VERSION)
LATEST_IMAGE_NAME = $(IMAGE_NAME):latest
STORAGE_TYPE = minio
LOCAL_STORAGE_PATH = generated_files

# 多架构支持
PLATFORMS = linux/amd64,linux/arm64

# 默认目标
help:
	@echo "可用的命令:"
	@echo "  install         - 安装项目依赖"
	@echo "  dev             - 启动开发服务器"
	@echo "  test            - 运行测试"
	@echo "  lint            - 代码检查"
	@echo "  format          - 代码格式化"
	@echo "  clean           - 清理临时文件"
	@echo "  docker-build    - 构建Docker镜像 (单架构)"
	@echo "  docker-buildx   - 构建Docker镜像 (多架构)"
	@echo "  docker-run      - 运行Docker容器"
	@echo "  docker-stop     - 停止Docker容器"
	@echo "  docker-push     - 推送镜像到仓库"
	@echo "  k8s-deploy      - 部署到Kubernetes"
	@echo "  k8s-undeploy    - 从Kubernetes卸载"
	@echo "  k8s-status      - 查看Kubernetes部署状态"
	@echo ""
	@echo "环境变量:"
	@echo "  REGISTRY        - Docker镜像仓库地址 (可选)"
	@echo "  VERSION         - 镜像版本 (默认: $(VERSION))"

# 安装依赖
install:
	@echo "创建虚拟环境并安装项目依赖..."
	uv venv
	uv pip install -e .
	uv pip install -e ".[dev]"

# 开发模式运行
dev:
	@echo "启动FastAPI服务器..."
	uv run uvicorn src.main:app --host 0.0.0.0 --port 8000 --reload

# 生产模式运行
prod:
	@echo "启动FastAPI服务器 (生产模式)..."
	uv run uvicorn src.main:app --host 0.0.0.0 --port 8000 --workers 4

# 直接运行main.py
run:
	@echo "直接运行main.py..."
	uv run python src/main.py

# 运行测试
test:
	@echo "运行测试..."
	uv run pytest tests/ -v

# 代码检查
lint:
	@echo "运行代码检查..."
	uv run flake8 src/ tests/
	uv run isort --check-only src/ tests/

# 代码格式化
format:
	@echo "格式化代码..."
	uv run black src/ tests/
	uv run isort src/ tests/

# 运行示例
example:
	@echo "运行使用示例..."
	uv run python example_usage.py

# 查看API文档
docs:
	@echo "API文档地址:"
	@echo "  Swagger UI: http://localhost:8000/docs"
	@echo "  ReDoc: http://localhost:8000/redoc"
openapi:
	@echo "生成 OpenAPI JSON (openapi.json)..."
	python tools/generate_openapi.py
	@echo "Swagger 文件将生成在 src/swagger/"

# 清理临时文件
clean:
	@echo "清理临时文件..."
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -delete
	find . -type d -name "*.egg-info" -exec rm -rf {} +
	rm -rf build/
	rm -rf dist/
	rm -rf .pytest_cache/
	rm -rf .coverage

# 构建Docker镜像 (单架构)
docker-build:
	@echo "构建Docker镜像 (单架构)..."
	docker build -t $(FULL_IMAGE_NAME) -t $(LATEST_IMAGE_NAME) .

# 构建Docker镜像 (多架构)
docker-buildx:
	@echo "构建Docker镜像 (多架构: $(PLATFORMS))..."
	@if ! docker buildx ls | grep -q multiarch; then \
		echo "创建多架构构建器..."; \
		docker buildx create --name multiarch --use; \
	fi
	docker buildx build \
		--platform $(PLATFORMS) \
		--tag $(FULL_IMAGE_NAME) \
		--tag $(LATEST_IMAGE_NAME) \
		--push .

# 构建Docker镜像 (多架构，本地)
docker-buildx-local:
	@echo "构建Docker镜像 (多架构，本地)..."
	@if ! docker buildx ls | grep -q multiarch; then \
		echo "创建多架构构建器..."; \
		docker buildx create --name multiarch --use; \
	fi
	docker buildx build \
		--platform $(PLATFORMS) \
		--tag $(FULL_IMAGE_NAME) \
		--tag $(LATEST_IMAGE_NAME) \
		--load .

# 推送镜像到仓库
docker-push:
	@echo "推送镜像到仓库..."
	docker push $(FULL_IMAGE_NAME)
	docker push $(LATEST_IMAGE_NAME)

# 运行Docker容器
docker-run:
	@echo "运行Docker容器..."
	docker run -d \
		--name $(PROJECT_NAME) \
		-p 8000:8000 \
		-v $(PWD)/$(LOCAL_STORAGE_PATH):/app/$(LOCAL_STORAGE_PATH) \
		--env-file .env \
		$(LATEST_IMAGE_NAME)

# 停止Docker容器
docker-stop:
	@echo "停止Docker容器..."
	docker stop $(PROJECT_NAME) || true
	docker rm $(PROJECT_NAME) || true

# 查看容器日志
docker-logs:
	docker logs -f $(PROJECT_NAME)

# 进入容器
docker-shell:
	docker exec -it $(PROJECT_NAME) /bin/bash

# 完整部署 (单架构)
deploy: docker-build docker-stop docker-run
	@echo "部署完成！"
	@echo "服务地址: http://localhost:8000"
	@echo "API文档: http://localhost:8000/docs"

# 完整部署 (多架构)
deploy-multiarch: docker-buildx-local docker-stop docker-run
	@echo "多架构部署完成！"
	@echo "服务地址: http://localhost:8000"
	@echo "API文档: http://localhost:8000/docs"

# Kubernetes部署
k8s-deploy:
	@echo "部署到Kubernetes..."
	@if [ ! -f deployment/deploy.sh ]; then \
		echo "❌ deployment/deploy.sh 不存在"; \
		exit 1; \
	fi
	cd deployment && ./deploy.sh dev

# Kubernetes完整部署
k8s-deploy-full:
	@echo "完整部署到Kubernetes (包含MinIO、Ingress、HPA)..."
	@if [ ! -f deployment/deploy.sh ]; then \
		echo "❌ deployment/deploy.sh 不存在"; \
		exit 1; \
	fi
	cd deployment && ./deploy.sh dev --with-minio --with-ingress --with-hpa

# Kubernetes卸载
k8s-undeploy:
	@echo "从Kubernetes卸载..."
	@if [ ! -f deployment/undeploy.sh ]; then \
		echo "❌ deployment/undeploy.sh 不存在"; \
		exit 1; \
	fi
	cd deployment && ./undeploy.sh

# 查看Kubernetes部署状态
k8s-status:
	@echo "Kubernetes部署状态:"
	@kubectl get pods -n ohc-account-invoice 2>/dev/null || echo "命名空间 ohc-account-invoice 不存在"
	@kubectl get services -n ohc-account-invoice 2>/dev/null || echo "服务不存在"
	@kubectl get ingress -n ohc-account-invoice 2>/dev/null || echo "Ingress不存在"
	@kubectl get hpa -n ohc-account-invoice 2>/dev/null || echo "HPA不存在"

# Kubernetes日志查看
k8s-logs:
	@echo "查看应用日志..."
	@kubectl logs -f deployment/ohc-account-invoice -n ohc-account-invoice

# Kubernetes进入Pod
k8s-shell:
	@echo "进入Pod..."
	@kubectl exec -it deployment/ohc-account-invoice -n ohc-account-invoice -- /bin/bash
