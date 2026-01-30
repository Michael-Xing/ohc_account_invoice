"""MinIO 存储服务（已迁移到 infrastructure 层）"""

import os
from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

from minio import Minio
from minio.error import S3Error
import importlib
import shutil
import base64

from src.config import settings


class StorageService(ABC):
    """存储服务抽象基类（保留以便未来扩展）"""

    @abstractmethod
    def save_file(self, file_path: Path, file_name: str, project_id: str = None, version: str = None) -> Tuple[bool, Optional[str], str]:
        pass

    @abstractmethod
    def get_file_url(self, file_name: str) -> Optional[str]:
        pass


class MinIOStorageService(StorageService):
    def __init__(self):
        if not settings.validate_minio_config():
            raise ValueError("MinIO配置不完整")

        self.minio_client = Minio(
            settings.minio_endpoint,
            access_key=settings.minio_access_key,
            secret_key=settings.minio_secret_key,
            secure=settings.minio_secure
        )
        self.bucket_name = settings.minio_bucket_name
        self._ensure_bucket_exists()

    def _ensure_bucket_exists(self):
        try:
            if not self.minio_client.bucket_exists(self.bucket_name):
                self.minio_client.make_bucket(self.bucket_name)
        except S3Error as e:
            raise ValueError(f"无法创建或访问存储桶: {str(e)}")

    def save_file(self, file_path: Path, file_name: str, project_id: str = None, version: str = None) -> Tuple[bool, Optional[str], str]:
        try:
            # 检查文件名是否已经包含时间戳（格式：YYYYMMDD-HHMMSS 或 YYYYMMDD_HHMMSS）
            import re
            name, ext = os.path.splitext(file_name)
            # 检查是否已经包含时间戳格式（YYYYMMDD-HHMMSS 或 YYYYMMDD_HHMMSS）
            timestamp_pattern = r'\d{8}[-_]\d{6}$'
            if re.search(timestamp_pattern, name):
                # 文件名已经包含时间戳，直接使用
                unique_name = file_name
            else:
                # 文件名不包含时间戳，添加时间戳
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                unique_name = f"{name}_{timestamp}{ext}"

            if project_id and version:
                object_key = f"{project_id}/{version}/{unique_name}"
            elif project_id:
                object_key = f"{project_id}/default/{unique_name}"
            else:
                object_key = f"default/default/{unique_name}"

            self.minio_client.fput_object(self.bucket_name, object_key, str(file_path))
            from datetime import timedelta
            file_url = self.minio_client.presigned_get_object(self.bucket_name, object_key, expires=timedelta(days=7))
            return True, file_url, "文件上传成功"
        except S3Error as e:
            return False, None, f"MinIO上传失败: {str(e)}"
        except Exception as e:
            return False, None, f"文件上传失败: {str(e)}"

    def get_file_url(self, file_name: str) -> Optional[str]:
        try:
            self.minio_client.stat_object(self.bucket_name, file_name)
            from datetime import timedelta
            return self.minio_client.presigned_get_object(self.bucket_name, file_name, expires=timedelta(days=7))
        except S3Error:
            return None


class StorageServiceFactory:
    @staticmethod
    def create_storage_service() -> StorageService:
        stype = getattr(settings, "storage_type", None)
        # 处理枚举类型：如果是枚举，使用 .value 获取字符串值；否则直接转换为字符串
        if stype is None:
            stype_val = "minio"
        elif hasattr(stype, "value"):
            # 枚举类型，使用 .value 获取实际值
            stype_val = str(stype.value).lower()
        else:
            stype_val = str(stype).lower()
        
        if stype_val == "local":
            return LocalStorageService()
        if stype_val in ("s3", "aws"):
            return S3StorageService()
        # default to MinIO
        return MinIOStorageService()


class LocalStorageService(StorageService):
    """本地文件系统存储实现"""
    def __init__(self):
        # ensure base path exists
        self.base_path = settings.get_local_storage_path()

    def save_file(self, file_path: Path, file_name: str, project_id: str = None, version: str = None) -> Tuple[bool, Optional[str], str]:
        try:
            target_dir = self.base_path / (project_id or "default") / (version or "default")
            target_dir.mkdir(parents=True, exist_ok=True)
            target_path = target_dir / file_name
            shutil.move(str(file_path), str(target_path))
            return True, str(target_path), "文件保存到本地成功"
        except Exception as e:
            return False, None, f"本地保存失败: {str(e)}"

    def get_file_url(self, file_name: str) -> Optional[str]:
        # return filesystem path for local files
        candidate = self.base_path / file_name
        if candidate.exists():
            return str(candidate)
        return None


class S3StorageService(StorageService):
    """AWS S3 存储实现（dynamic boto3 import）"""
    def __init__(self):
        if not settings.validate_s3_config():
            raise ValueError("S3 配置不完整")
        try:
            boto3 = importlib.import_module("boto3")
        except Exception as e:
            raise ImportError("boto3 未安装，S3 存储不可用") from e

        self.client = boto3.client(
            "s3",
            aws_access_key_id=settings.aws_access_key_id,
            aws_secret_access_key=settings.aws_secret_access_key,
            region_name=getattr(settings, "aws_region", None),
        )
        self.bucket = settings.aws_bucket_name

    def save_file(self, file_path: Path, file_name: str, project_id: str = None, version: str = None) -> Tuple[bool, Optional[str], str]:
        try:
            import time
            import re
            # 检查文件名是否已经包含时间戳（格式：YYYYMMDD-HHMMSS 或 YYYYMMDD_HHMMSS）
            name, ext = os.path.splitext(file_name)
            # 检查是否已经包含时间戳格式（YYYYMMDD-HHMMSS 或 YYYYMMDD_HHMMSS）
            timestamp_pattern = r'\d{8}[-_]\d{6}$'
            if re.search(timestamp_pattern, name):
                # 文件名已经包含时间戳，直接使用
                unique_name = file_name
            else:
                # 文件名不包含时间戳，添加时间戳
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                unique_name = f"{name}_{timestamp}{ext}"
            if project_id and version:
                key = f"{project_id}/{version}/{unique_name}"
            elif project_id:
                key = f"{project_id}/default/{unique_name}"
            else:
                key = f"default/default/{unique_name}"

            self.client.upload_file(str(file_path), self.bucket, key)
            # presigned url 7 days
            url = self.client.generate_presigned_url(
                "get_object",
                Params={"Bucket": self.bucket, "Key": key},
                ExpiresIn=7 * 24 * 3600,
            )
            return True, url, "文件上传到 S3 成功"
        except Exception as e:
            return False, None, f"S3 上传失败: {str(e)}"

    def get_file_url(self, file_name: str) -> Optional[str]:
        try:
            url = self.client.generate_presigned_url(
                "get_object",
                Params={"Bucket": self.bucket, "Key": file_name},
                ExpiresIn=7 * 24 * 3600,
            )
            return url
        except Exception:
            return None


