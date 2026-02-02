from typing import Optional
from pydantic import BaseModel, Field


class BaseTemplateParameters(BaseModel):
    project_name: Optional[str] = Field(default=None, description="项目名称")
    version: Optional[str] = Field(default=None, description="版本号")
    date: Optional[str] = Field(default=None, description="日期")
    author: Optional[str] = Field(default="OHC账票AI助手", description="作者")
    document_type: Optional[str] = Field(default=None, description="文档类型")
    department: Optional[str] = Field(default=None, description="部门")
    language: Optional[str] = Field(default=None, description="语言代码 (zh/ja/en)，如果不指定则使用默认模板")


class DHFIndexParameters(BaseTemplateParameters):
    """文件･图纸一览（`static/templates/excel/zh/文件･图纸一览.xlsx`）参数"""
    theme_no: str = Field(..., description="项目NO，填入C3单元格")
    theme_name: str = Field(..., description="项目名称，填入C4单元格")
    product_model_name: str = Field(..., description="商品型式名，填入C5单元格")
    sales_name: str = Field(..., description="贩卖名称，填入C6单元格")
    stage: str = Field(..., description="设计阶段，填入C6单元格")

class PTFIndexParameters(BaseTemplateParameters):
    """PTF INDEX（`static/templates/excel/zh/PTF INDEX.xlsx`）参数"""
    pass


class ESIndividualTestSpecParameters(BaseTemplateParameters):
    """个别试验要项书（`static/templates/excel/zh/个别试验要项书.xlsx`）参数"""

    test_item: Optional[str] = None


class ESIndividualTestResultParameters(BaseTemplateParameters):
    test_item: Optional[str] = None
    test_result: Optional[str] = None
    tester: Optional[str] = None


class PPIndividualTestResultParameters(BaseTemplateParameters):
    test_item: Optional[str] = None
    test_result: Optional[str] = None


class ESVerificationPlanParameters(BaseTemplateParameters):
    verification_purpose: Optional[str] = None


class ESVerificationResultParameters(BaseTemplateParameters):
    verification_result: Optional[str] = None


class PPVerificationPlanParameters(BaseTemplateParameters):
    verification_purpose: Optional[str] = None


class PPVerificationResultParameters(BaseTemplateParameters):
    verification_result: Optional[str] = None


class BasicSpecificationParameters(BaseTemplateParameters):
    overview: Optional[str] = None
    technical_requirements: Optional[str] = None
    acceptance_criteria: Optional[str] = None


class PPIndividualTestSpecParameters(BaseTemplateParameters):
    test_purpose: Optional[str] = None


class FollowUpDRMinutesParameters(BaseTemplateParameters):
    meeting_date: Optional[str] = None
    meeting_location: Optional[str] = None


class LabelingSpecificationParameters(BaseTemplateParameters):
    """标签仕样书-仕样确认书（`static/templates/excel/zh/标签仕样书-仕样确认书.xlsx`）参数"""
    theme_no: str = Field(..., description="项目NO，填入D5单元格")
    theme_name: str = Field(..., description="项目名称，填入K5单元格")
    product_model_name: str = Field(..., description="商品型式名，填入D7单元格")

class ProductEnvironmentAssessmentParameters(BaseTemplateParameters):
    """产品环境评估要项书/结果书参数"""
    theme_no: str = Field(..., description="项目NO，拼接到B5单元格内容后面")
    theme_name: Optional[str] = Field(default=None, description="商品类别（已废弃，不再使用）")
    product_model: str = Field(..., description="商品型号，根据'/'分割，填充到22行的D～H列合并单元格")
    product_model_name: str = Field(..., description="商品型号名，拼接到B7单元格内容后面")
    product_name: str = Field(..., description="商品名，拼接到I5单元格内容后面")
    production_area: str = Field(..., description="生产地，拼接到I7单元格内容后面")
    sales_name: str = Field(..., description="贩卖名称，根据'/'分割，填充到22行的I～J列合并单元格")
    target_area: str = Field(..., description="贩卖国家，填充到22行的K～M列合并单元格")
    remarks: Optional[str] = Field(default=None, description="备注（可选）")
    eta_schedule: Optional[str] = Field(default=None, description="ETA预定日志（可选）")


class ExistingProductComparisonParameters(BaseTemplateParameters):
    comparison_products: Optional[str] = None
    comparison_results: Optional[str] = None


class PackagingDesignSpecificationParameters(BaseTemplateParameters):
    """包装设计仕样书（`static/templates/excel/zh/包装设计仕样书.xlsx`）参数"""
    theme_no: str = Field(..., description="项目NO，填入C21单元格")
    theme_name: str = Field(..., description="项目名称，填入E21单元格")
    product_model_name: str = Field(..., description="商品型式名，填入L21单元格")
    sales_name: str = Field(..., description="贩卖名称，填入C23单元格")


class UserManualSpecificationParameters(BaseTemplateParameters):
    """使用说明书仕样书（`static/templates/excel/zh/使用说明书仕样书.xlsx`）参数"""
    theme_no: str = Field(..., description="项目NO，填入B19单元格")
    theme_name: str = Field(..., description="项目名称，填入D19单元格")
    product_model_name: str = Field(..., description="商品型式名，填入J19单元格")
    sales_name: str = Field(..., description="贩卖名称，填入B21单元格")
    file_type: str = Field(..., description="文件类型，按需填写")
    name: str = Field(..., description="名称，按需填写")
    version: str = Field(..., description="版本，按需填写")


class ProjectPlanParameters(BaseTemplateParameters):
    project_scope: Optional[str] = None
    project_timeline: Optional[str] = None

