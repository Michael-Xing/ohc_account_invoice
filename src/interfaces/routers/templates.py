from fastapi import APIRouter, HTTPException
from typing import Dict, Optional

from src.application.template_service import get_supported_templates, get_template_info
from src.interfaces.schemas import TemplateInfoResponse

router = APIRouter(prefix="", tags=["templates"])


@router.get("/templates", summary="获取模板列表", description="获取所有支持的模板列表")
async def list_templates(language: Optional[str] = None):
    """
    获取模板列表
    
    Args:
        language: 可选的语言代码 (zh/ja/en)
            - 如果指定：返回 {template_name: display_name} 格式，只包含该语言的显示名称
            - 如果未指定：返回 {template_name: description} 格式，包含描述和所有语言的显示名称
    """
    return get_supported_templates(language)


@router.get("/templates/{template_name}", response_model=TemplateInfoResponse, summary="获取模板信息", description="获取指定模板的详细信息")
async def get_template_info_route(template_name: str, language: Optional[str] = None):
    info = get_template_info(template_name, language)
    if not info:
        raise HTTPException(status_code=404, detail=f"模板 '{template_name}' 不存在")
    return TemplateInfoResponse(**info)


