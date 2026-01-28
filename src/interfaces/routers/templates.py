from fastapi import APIRouter, HTTPException
from typing import Dict

from src.application.template_service import get_supported_templates, get_template_info
from src.interfaces.schemas import TemplateInfoResponse

router = APIRouter(prefix="", tags=["templates"])


@router.get("/templates", response_model=Dict[str, str], summary="获取模板列表", description="获取所有支持的模板列表")
async def list_templates():
    return get_supported_templates()


@router.get("/templates/{template_name}", response_model=TemplateInfoResponse, summary="获取模板信息", description="获取指定模板的详细信息")
async def get_template_info_route(template_name: str):
    info = get_template_info(template_name)
    if not info:
        raise HTTPException(status_code=404, detail=f"模板 '{template_name}' 不存在")
    return TemplateInfoResponse(**info)


