"""配置管理模块"""

from enum import Enum
from pathlib import Path
from typing import Optional

import os
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class StorageType(str, Enum):
    """存储类型枚举"""
    MINIO = "minio"
    LOCAL = "local"
    S3 = "s3"


env_file_value = None if os.environ.get("SKIP_DOTENV") or os.environ.get("SKIP_INFRA_INIT") else ".env"


class Settings(BaseSettings):
    """应用配置"""
    model_config = SettingsConfigDict(
        env_file=env_file_value,
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore"
    )
    
    # 存储配置（固定为 MinIO）
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
    # AWS S3 配置（移除——不再支持 S3）
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
    
    # 模板配置（改为 static 下的 templates）
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


# 全局配置实例
try:
    settings = Settings()
except PermissionError:
    # Fallback: create settings without reading .env by forcing env_file=None
    temp_config = SettingsConfigDict(env_file=None)
    Settings.model_config = temp_config
    settings = Settings()
