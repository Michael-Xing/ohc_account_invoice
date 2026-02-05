from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field, ConfigDict


class GenerateDocumentRequest(BaseModel):
    """生成文档请求模型"""
    template_name: str = Field(..., description="模板名称")
    parameters: Dict[str, Any] = Field(..., description="模板参数")
    language: Optional[str] = Field(default=None, description="语言代码 (zh/ja/en)，如果不指定则使用默认模板")

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "template_name": "DHF_INDEX",
                "parameters": {
                    "project_number": "OHC项目",
                    "version": "1.0",
                    "date": "2025-01-22",
                    "author": "张三",
                    "document_type": "设计文档",
                    "department": "研发部"
                },
                "language": "zh"
            }
        }
    )


class GenerateDocumentResponse(BaseModel):
    """生成文档响应模型"""
    success: bool = Field(..., description="是否成功")
    message: str = Field(..., description="响应消息")
    file_name: Optional[str] = Field(None, description="生成的文件名")
    file_url: Optional[str] = Field(None, description="文件下载链接（远程存储时）或本地文件路径（本地存储时）")
    storage_type: Optional[str] = Field(None, description="存储类型")
    project_id: Optional[str] = Field(None, description="项目编号")
    version: Optional[str] = Field(None, description="版本号")

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "success": True,
                "message": "文档生成成功",
                "file_name": "OHC项目_v1.0_20250122_DHF_INDEX.xlsx",
                "file_url": "/path/to/file.xlsx",
                "storage_type": "local",
                "project_id": "OHC项目",
                "version": "1.0"
            }
        }
    )


class TemplateInfoResponse(BaseModel):
    """模板信息响应模型"""
    name: str = Field(..., description="模板名称")
    display_name: str = Field(..., description="显示名称")
    description: Optional[str] = Field(None, description="模板描述")
    available_formats: List[str] = Field(..., description="可用格式")
    filler_strategy: str = Field(..., description="填充策略")
    features: List[str] = Field(..., description="特性列表")
    available_languages: List[str] = Field(default_factory=list, description="可用语言列表")
    display_names: Optional[Dict[str, str]] = Field(None, description="各语言的显示名称")


class ServiceConfigResponse(BaseModel):
    """服务配置响应模型"""
    storage_type: str = Field(..., description="存储类型")
    template_base_path: str = Field(..., description="模板基础路径")
    supported_templates_count: int = Field(..., description="支持的模板数量")
    app_name: str = Field(..., description="应用名称")
    app_version: str = Field(..., description="应用版本")


class HealthCheckResponse(BaseModel):
    """健康检查响应模型"""
    status: str = Field(..., description="服务状态")
    service: str = Field(..., description="服务名称")
    timestamp: str = Field(..., description="检查时间")


# 导出模板参数模型（位于同包的 templates 模块）
from .templates import (
    BaseTemplateParameters,
    DHFIndexParameters,
    PTFIndexParameters,
    ESIndividualTestSpecParameters,
    ESIndividualTestResultParameters,
    PPIndividualTestResultParameters,
    ESVerificationPlanParameters,
    ESVerificationResultParameters,
    PPVerificationPlanParameters,
    PPVerificationResultParameters,
    BasicSpecificationParameters,
    PPIndividualTestSpecParameters,
    FollowUpDRMinutesParameters,
    LabelingSpecificationParameters,
    ProductEnvironmentAssessmentParameters,
    ExistingProductComparisonParameters,
    PackagingDesignSpecificationParameters,
    UserManualSpecificationParameters,
    ProjectPlanParameters,
)


