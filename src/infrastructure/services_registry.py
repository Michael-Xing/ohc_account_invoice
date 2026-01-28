"""服务注册表，用于创建共享服务实例。"""
import os
from src.config import settings
from src.domain.template_filler_service import TemplateService
from src.infrastructure.storage_service import StorageServiceFactory

# 在此处实例化应用级别的单例服务
template_service = TemplateService()

# 在 CI 或受限环境中，可能需要跳过初始化会进行网络访问的外部资源（例如 MinIO）。
# 当环境变量 SKIP_INFRA_INIT 设置为 "1"/"true"/"yes" 时，将跳过 storage_service 的初始化，
# 以避免在模块导入时发生网络调用或抛出配置相关的错误。
if os.environ.get("SKIP_INFRA_INIT", "").lower() in ("1", "true", "yes"):
    storage_service = None
else:
    storage_service = StorageServiceFactory.create_storage_service()


