#!/bin/bash

# OHC账票生成服务 Kubernetes部署脚本
# 使用方法: ./deploy.sh [环境] [选项]
# 环境: dev, staging, prod
# 选项: --with-minio, --with-ingress, --with-hpa

set -e

# 默认配置
ENVIRONMENT=${1:-dev}
NAMESPACE="ohc-account-invoice"
IMAGE_TAG="1.0.0"
REGISTRY=""

# 解析参数
WITH_MINIO=false
WITH_INGRESS=false
WITH_HPA=false

for arg in "$@"; do
    case $arg in
        --with-minio)
            WITH_MINIO=true
            ;;
        --with-ingress)
            WITH_INGRESS=true
            ;;
        --with-hpa)
            WITH_HPA=true
            ;;
    esac
done

echo "🚀 开始部署 OHC账票生成服务"
echo "环境: $ENVIRONMENT"
echo "命名空间: $NAMESPACE"
echo "镜像标签: $IMAGE_TAG"

# 检查kubectl是否可用
if ! command -v kubectl &> /dev/null; then
    echo "❌ kubectl 未安装或不在PATH中"
    exit 1
fi

# 检查集群连接
if ! kubectl cluster-info &> /dev/null; then
    echo "❌ 无法连接到Kubernetes集群"
    exit 1
fi

echo "✅ Kubernetes集群连接正常"

# 创建命名空间
echo "📦 创建命名空间..."
kubectl apply -f namespace.yaml

# 应用ConfigMap
echo "⚙️  应用配置..."
kubectl apply -f configmap.yaml

# 应用Secret
echo "🔐 应用密钥配置..."
kubectl apply -f secret.yaml

# 部署MinIO (如果启用)
if [ "$WITH_MINIO" = true ]; then
    echo "🗄️  部署MinIO存储..."
    kubectl apply -f minio.yaml
    echo "⏳ 等待MinIO启动..."
    kubectl wait --for=condition=available --timeout=300s deployment/minio -n $NAMESPACE
fi

# 部署应用
echo "🚀 部署应用..."
kubectl apply -f deployment.yaml

# 应用Service
echo "🌐 应用服务配置..."
kubectl apply -f service.yaml

# 应用Ingress (如果启用)
if [ "$WITH_INGRESS" = true ]; then
    echo "🔗 应用Ingress配置..."
    kubectl apply -f ingress.yaml
fi

# 应用HPA (如果启用)
if [ "$WITH_HPA" = true ]; then
    echo "📈 应用HPA配置..."
    kubectl apply -f hpa.yaml
fi

# 等待部署完成
echo "⏳ 等待部署完成..."
kubectl wait --for=condition=available --timeout=300s deployment/ohc-account-invoice -n $NAMESPACE

# 显示部署状态
echo "📊 部署状态:"
kubectl get pods -n $NAMESPACE
kubectl get services -n $NAMESPACE

# 显示访问信息
echo ""
echo "🎉 部署完成!"
echo ""
echo "📋 访问信息:"
echo "  集群内访问: http://ohc-account-invoice-service.$NAMESPACE.svc.cluster.local:8000"
echo "  API文档: http://ohc-account-invoice-service.$NAMESPACE.svc.cluster.local:8000/docs"
echo "  健康检查: http://ohc-account-invoice-service.$NAMESPACE.svc.cluster.local:8000/health"

if [ "$WITH_INGRESS" = true ]; then
    echo ""
    echo "🌐 外部访问:"
    echo "  请根据Ingress配置的域名访问服务"
fi

if [ "$WITH_MINIO" = true ]; then
    echo ""
    echo "🗄️  MinIO访问:"
    echo "  管理界面: http://minio-console-service.$NAMESPACE.svc.cluster.local:9001"
    echo "  用户名: minioadmin"
    echo "  密码: minioadmin"
fi

echo ""
echo "🔧 常用命令:"
echo "  查看Pod状态: kubectl get pods -n $NAMESPACE"
echo "  查看日志: kubectl logs -f deployment/ohc-account-invoice -n $NAMESPACE"
echo "  进入Pod: kubectl exec -it deployment/ohc-account-invoice -n $NAMESPACE -- /bin/bash"
echo "  删除部署: kubectl delete namespace $NAMESPACE"
