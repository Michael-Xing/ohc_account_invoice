from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from datetime import datetime
from typing import Dict

from src.infrastructure.services_registry import template_service, storage_service
from src.interfaces.schemas import HealthCheckResponse, ServiceConfigResponse
from src.config import settings

router = APIRouter(prefix="", tags=["system"])


@router.get("/health", response_model=HealthCheckResponse, summary="健康检查", description="检查服务健康状态")
async def health_check():
    return HealthCheckResponse(
        status="healthy",
        service=settings.app_name,
        timestamp=datetime.now().isoformat()
    )


@router.get("/config", response_model=ServiceConfigResponse, summary="获取服务配置", description="获取服务配置信息")
async def get_service_config():
    return ServiceConfigResponse(
        storage_type=settings.storage_type if not hasattr(settings, "storage_type") else settings.storage_type.value,
        template_base_path=settings.template_base_path,
        supported_templates_count=len(template_service.get_supported_templates(None)),
        app_name=settings.app_name,
        app_version=settings.app_version
    )


@router.get("/download/{filename}", summary="下载文件", description="下载生成的文件（仅本地存储时可用）")
async def download_file(filename: str):
    # 处理枚举类型：如果是枚举，使用 .value 获取字符串值；否则直接转换为字符串
    storage_type_val = None
    if hasattr(settings, "storage_type") and settings.storage_type is not None:
        if hasattr(settings.storage_type, "value"):
            storage_type_val = str(settings.storage_type.value).lower()
        else:
            storage_type_val = str(settings.storage_type).lower()
    
    if storage_type_val != "local":
        raise HTTPException(status_code=400, detail="文件下载仅支持本地存储模式")

    file_path = settings.get_local_storage_path() / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/octet-stream'
    )


@router.get("/", summary="根路径", description="返回API基本信息")
async def root():
    """根路径"""
    return {
        "message": "OHC账票生成API服务",
        "version": settings.app_version,
        "docs": "/docs",
        "redoc": "/redoc"
    }


