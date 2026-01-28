#!/bin/bash

# OHC账票生成服务 Kubernetes卸载脚本
# 使用方法: ./undeploy.sh [选项]
# 选项: --with-minio, --force

set -e

NAMESPACE="ohc-account-invoice"
WITH_MINIO=false
FORCE=false

# 解析参数
for arg in "$@"; do
    case $arg in
        --with-minio)
            WITH_MINIO=true
            ;;
        --force)
            FORCE=true
            ;;
    esac
done

echo "🗑️  开始卸载 OHC账票生成服务"
echo "命名空间: $NAMESPACE"

# 检查kubectl是否可用
if ! command -v kubectl &> /dev/null; then
    echo "❌ kubectl 未安装或不在PATH中"
    exit 1
fi

# 检查命名空间是否存在
if ! kubectl get namespace $NAMESPACE &> /dev/null; then
    echo "❌ 命名空间 $NAMESPACE 不存在"
    exit 1
fi

# 确认删除
if [ "$FORCE" != true ]; then
    echo "⚠️  这将删除命名空间 $NAMESPACE 中的所有资源"
    read -p "确认删除? (y/N): " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        echo "❌ 取消删除"
        exit 1
    fi
fi

# 删除HPA
echo "📈 删除HPA..."
kubectl delete hpa ohc-account-invoice-hpa -n $NAMESPACE --ignore-not-found=true

# 删除Ingress
echo "🔗 删除Ingress..."
kubectl delete ingress ohc-account-invoice-ingress -n $NAMESPACE --ignore-not-found=true
kubectl delete ingress ohc-account-invoice-ingress-tls -n $NAMESPACE --ignore-not-found=true

# 删除Service
echo "🌐 删除Service..."
kubectl delete service ohc-account-invoice-service -n $NAMESPACE --ignore-not-found=true
kubectl delete service ohc-account-invoice-nodeport -n $NAMESPACE --ignore-not-found=true
kubectl delete service ohc-account-invoice-loadbalancer -n $NAMESPACE --ignore-not-found=true

# 删除Deployment
echo "🚀 删除Deployment..."
kubectl delete deployment ohc-account-invoice -n $NAMESPACE --ignore-not-found=true

# 删除MinIO相关资源 (如果启用)
if [ "$WITH_MINIO" = true ]; then
    echo "🗄️  删除MinIO..."
    kubectl delete deployment minio -n $NAMESPACE --ignore-not-found=true
    kubectl delete service minio-service -n $NAMESPACE --ignore-not-found=true
    kubectl delete service minio-console-service -n $NAMESPACE --ignore-not-found=true
    kubectl delete pvc minio-pvc -n $NAMESPACE --ignore-not-found=true
fi

# 删除ConfigMap和Secret
echo "⚙️  删除配置..."
kubectl delete configmap ohc-account-invoice-config -n $NAMESPACE --ignore-not-found=true
kubectl delete secret ohc-account-invoice-secret -n $NAMESPACE --ignore-not-found=true

# 删除命名空间
echo "📦 删除命名空间..."
kubectl delete namespace $NAMESPACE --ignore-not-found=true

echo "✅ 卸载完成!"
echo ""
echo "🔍 验证删除结果:"
kubectl get namespace $NAMESPACE --ignore-not-found=true

if [ "$WITH_MINIO" = true ]; then
    echo ""
    echo "⚠️  注意: MinIO的持久化数据可能仍然存在"
    echo "如需完全清理，请手动删除相关的PV和PVC"
fi
