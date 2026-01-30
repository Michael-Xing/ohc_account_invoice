"""配置管理模块"""

from enum import Enum
from pathlib import Path
from typing import Optional, Dict, Any
import os

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class StorageType(str, Enum):
    """存储类型枚举"""
    MINIO = "minio"
    LOCAL = "local"
    S3 = "s3"


def load_toml_config(config_path: Optional[str] = None) -> Dict[str, Any]:
    """
    从 TOML 文件加载配置
    
    Args:
        config_path: TOML 配置文件路径，如果为 None 则尝试查找 config.toml
        
    Returns:
        配置字典
    """
    
    if config_path is None:
        # 尝试多个可能的路径（按优先级）
        possible_paths = [
            Path.cwd() / "config" / "config.toml",  # 优先查找 config/ 目录
            Path.cwd() / "config.toml",  # 根目录下的 config.toml
            Path(__file__).parent.parent.parent / "config" / "config.toml",  # 项目根目录的 config/ 目录
            Path(__file__).parent.parent.parent / "config.toml",  # 项目根目录下的 config.toml
        ]
        for path in possible_paths:
            if path.exists():
                config_path = str(path)
                break
        
        if config_path is None:
            return {}
    
    config_file = Path(config_path)
    if not config_file.exists():
        return {}
    
    # 读取并解析 TOML 文件
    try:
        # 获取 TOML 解析库
        try:
            import tomllib as toml_parser  # Python 3.11+
        except ImportError:
            import tomli as toml_parser  # type: ignore  # 兼容旧版本，可选依赖
        
        with open(config_file, "rb") as f:
            return toml_parser.load(f)
    except Exception as e:
        # 如果解析失败，返回空字典
        print(f"警告: 无法解析 TOML 配置文件 {config_path}: {e}")
        return {}


def get_toml_file_path() -> Optional[str]:
    """获取 TOML 配置文件路径"""
    # 优先使用环境变量指定的路径
    config_path = os.environ.get("CONFIG_FILE")
    if config_path and Path(config_path).exists():
        return config_path
    
    # 尝试查找 config.toml（按优先级）
    possible_paths = [
        Path.cwd() / "config" / "config.toml",  # 优先查找 config/ 目录
        Path.cwd() / "config.toml",  # 根目录下的 config.toml
        Path(__file__).parent.parent.parent / "config" / "config.toml",  # 项目根目录的 config/ 目录
        Path(__file__).parent.parent.parent / "config.toml",  # 项目根目录下的 config.toml
    ]
    for path in possible_paths:
        if path.exists():
            return str(path)
    
    return None


class Settings(BaseSettings):
    """应用配置（从 config.toml 文件加载）"""
    model_config = SettingsConfigDict(
        env_file=None,  # 不再使用 .env 文件，只从 TOML 文件加载
        case_sensitive=False,
        extra="ignore"
    )
    
    def __init__(self, **kwargs):
        # 从 TOML 文件加载配置
        toml_data = load_toml_config()
        if toml_data:
            # 将 TOML 配置扁平化并合并到 kwargs
            flattened = _flatten_toml_config(toml_data)
            kwargs.update(flattened)
        super().__init__(**kwargs)
    
    # 存储配置
    storage_type: StorageType = Field(default=StorageType.MINIO, description="存储类型")
    
    # MinIO 配置
    minio_endpoint: Optional[str] = Field(default=None, description="MinIO服务端点")
    minio_access_key: Optional[str] = Field(default=None, description="MinIO访问密钥")
    minio_secret_key: Optional[str] = Field(default=None, description="MinIO秘密密钥")
    minio_bucket_name: Optional[str] = Field(default=None, description="MinIO存储桶名称")
    minio_secure: bool = Field(default=True, description="是否使用HTTPS连接MinIO")
    
    # 本地存储配置
    local_storage_path: str = Field(default="generated_files", description="本地存储基础路径")
    
    # AWS S3 配置
    aws_access_key_id: Optional[str] = Field(default=None, description="AWS访问密钥ID")
    aws_secret_access_key: Optional[str] = Field(default=None, description="AWS秘密访问密钥")
    aws_bucket_name: Optional[str] = Field(default=None, description="S3存储桶名称")
    aws_region: str = Field(default="us-east-1", description="AWS区域")
    
    # FastAPI应用配置
    app_name: str = Field(default="OHC账票生成服务", description="应用名称")
    app_version: str = Field(default="1.0.0", description="应用版本")
    app_description: str = Field(default="OHC账票生成API服务", description="应用描述")
    host: str = Field(default="0.0.0.0", description="服务器主机地址")
    port: int = Field(default=8000, description="服务器端口")
    debug: bool = Field(default=False, description="调试模式")
    reload: bool = Field(default=False, description="自动重载")
    
    # 模板配置
    template_base_path: str = Field(default="static/templates", description="模板基础路径")
    
    # 文件名配置
    filename_include_timestamp: bool = Field(default=True, description="文件名是否包含时间戳")
    filename_max_length: int = Field(default=200, description="文件名最大长度")
    
    # Sentry / monitoring (optional)
    sentry_dsn: Optional[str] = Field(default=None, description="Sentry DSN (optional)")
    sentry_environment: Optional[str] = Field(default=None, description="Sentry environment name")
    
    def get_template_path(self, template_type: str) -> Path:
        """获取模板路径"""
        return Path(self.template_base_path) / template_type
    
    def get_local_storage_path(self) -> Path:
        """返回本地存储路径（Path），并在必要时创建父目录。"""
        p = Path(self.local_storage_path)
        p.mkdir(parents=True, exist_ok=True)
        return p
    
    def validate_s3_config(self) -> bool:
        """验证 S3 配置是否完整"""
        required = [
            self.aws_access_key_id,
            self.aws_secret_access_key,
            self.aws_bucket_name,
        ]
        return all(field is not None for field in required)
    
    def validate_minio_config(self) -> bool:
        """验证 MinIO 配置是否完整（MinIO 为强制）"""
        required_fields = [
            self.minio_endpoint,
            self.minio_access_key,
            self.minio_secret_key,
            self.minio_bucket_name
        ]
        return all(field is not None for field in required_fields)


def _flatten_toml_config(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    将嵌套的 TOML 配置扁平化并映射到 Settings 字段名
    
    TOML 结构映射规则:
    - app.* -> app_*
    - storage.type -> storage_type
    - storage.minio.* -> minio_*
    - storage.local.path -> local_storage_path
    - storage.s3.* -> aws_* (特殊映射)
    - templates.base_path -> template_base_path
    - files.* -> filename_*
    - monitoring.* -> sentry_*
    """
    result = {}
    
    # 应用配置
    if "app" in data:
        app_config = data["app"]
        # app 配置的特殊映射
        app_mapping = {
            "name": "app_name",
            "version": "app_version",
            "description": "app_description",
            "host": "host",
            "port": "port",
            "debug": "debug",
            "reload": "reload",
        }
        for key, value in app_config.items():
            result_key = app_mapping.get(key, f"app_{key}")
            result[result_key] = value
    
    # 存储配置
    if "storage" in data:
        storage_config = data["storage"]
        # storage.type -> storage_type
        if "type" in storage_config:
            result["storage_type"] = storage_config["type"]
        
        # MinIO 配置
        if "minio" in storage_config:
            minio_config = storage_config["minio"]
            for key, value in minio_config.items():
                result[f"minio_{key}"] = value
        
        # 本地存储配置
        if "local" in storage_config:
            local_config = storage_config["local"]
            if "path" in local_config:
                result["local_storage_path"] = local_config["path"]
        
        # S3 配置（映射到 aws_ 前缀）
        if "s3" in storage_config:
            s3_config = storage_config["s3"]
            mapping = {
                "access_key_id": "aws_access_key_id",
                "secret_access_key": "aws_secret_access_key",
                "bucket_name": "aws_bucket_name",
                "region": "aws_region",
            }
            for key, value in s3_config.items():
                result_key = mapping.get(key, f"aws_{key}")
                result[result_key] = value
    
    # 模板配置
    if "templates" in data:
        templates_config = data["templates"]
        if "base_path" in templates_config:
            result["template_base_path"] = templates_config["base_path"]
    
    # 文件配置
    if "files" in data:
        files_config = data["files"]
        if "include_timestamp" in files_config:
            result["filename_include_timestamp"] = files_config["include_timestamp"]
        if "max_length" in files_config:
            result["filename_max_length"] = files_config["max_length"]
    
    # 监控配置
    if "monitoring" in data:
        monitoring_config = data["monitoring"]
        if "sentry_dsn" in monitoring_config:
            result["sentry_dsn"] = monitoring_config["sentry_dsn"]
        if "sentry_environment" in monitoring_config:
            result["sentry_environment"] = monitoring_config["sentry_environment"]
    
    return result


# 全局配置实例（从 config.toml 文件加载）
settings = Settings()
