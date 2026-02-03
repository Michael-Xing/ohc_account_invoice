from typing import Optional, List
from pydantic import BaseModel, Field


class BaseTemplateParameters(BaseModel):
    project_name: Optional[str] = Field(default=None, description="项目名称")
    version: Optional[str] = Field(default=None, description="版本号")
    date: Optional[str] = Field(default=None, description="日期")
    author: Optional[str] = Field(default="OHC账票AI助手", description="作者")
    document_type: Optional[str] = Field(default=None, description="文档类型")
    department: Optional[str] = Field(default=None, description="部门")
    language: Optional[str] = Field(default=None, description="语言代码 (zh/ja/en)，如果不指定则使用默认模板")


class FileListItem(BaseModel):
    """文件列表项"""
    number: str = Field(..., description="文件编号")
    file_name: str = Field(..., description="文件名称")
    stage: str = Field(..., description="阶段")


class DHFIndexParameters(BaseTemplateParameters):
    """DHF INDEX参数"""
    theme_no: str = Field(..., description="项目NO，填充到C3单元格")
    theme_name: str = Field(..., description="商品类别，填充到C4单元格")
    product_model: str = Field(..., description="商品型号，根据'/'分割，填充到C5单元格")
    sales_name: str = Field(..., description="贩卖名称，根据'/'分割，填充到C6单元格")
    stage: str = Field(..., description="阶段，拼接到C7单元格内容的后面")
    product_name: str = Field(..., description="商品名")
    file_list: List[FileListItem] = Field(..., description="文件列表")


class PTFIndexParameters(BaseTemplateParameters):
    """PTF INDEX参数"""
    pass


class ESIndividualTestSpecParameters(BaseTemplateParameters):
    """ES个别试验要项书参数"""
    test_item: Optional[str] = Field(default=None, description="试验项目")


class ESIndividualTestResultParameters(BaseTemplateParameters):
    """ES个别试验结果书参数"""
    test_item: Optional[str] = Field(default=None, description="试验项目")
    test_result: Optional[str] = Field(default=None, description="试验结果")
    tester: Optional[str] = Field(default=None, description="试验者")


class PPIndividualTestResultParameters(BaseTemplateParameters):
    """PP个别试验结果书参数"""
    test_item: Optional[str] = Field(default=None, description="试验项目")
    test_result: Optional[str] = Field(default=None, description="试验结果")


class ESVerificationPlanParameters(BaseTemplateParameters):
    """ES验证计划书参数"""
    verification_purpose: Optional[str] = Field(default=None, description="验证目的")


class ESVerificationResultParameters(BaseTemplateParameters):
    """ES验证结果书参数"""
    verification_result: Optional[str] = Field(default=None, description="验证结果")


class PPVerificationPlanParameters(BaseTemplateParameters):
    """PP验证计划书参数"""
    verification_purpose: Optional[str] = Field(default=None, description="验证目的")


class PPVerificationResultParameters(BaseTemplateParameters):
    """PP验证结果书参数"""
    verification_result: Optional[str] = Field(default=None, description="验证结果")


class BasicSpecificationParameters(BaseTemplateParameters):
    """基本规格书参数"""
    overview: Optional[str] = Field(default=None, description="概述")
    technical_requirements: Optional[str] = Field(default=None, description="技术要求")
    acceptance_criteria: Optional[str] = Field(default=None, description="验收标准")


class PPIndividualTestSpecParameters(BaseTemplateParameters):
    """PP个别试验要项书参数"""
    test_purpose: Optional[str] = Field(default=None, description="试验目的")


class FollowUpDRMinutesParameters(BaseTemplateParameters):
    """跟进DR会议记录参数"""
    meeting_date: Optional[str] = Field(default=None, description="会议日期")
    meeting_location: Optional[str] = Field(default=None, description="会议地点")


class LabelingSpecificationParameters(BaseTemplateParameters):
    """标签规格书参数"""
    product_name: Optional[str] = Field(default=None, description="产品名称")


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
    """与现有产品对比表参数"""
    comparison_products: Optional[str] = Field(default=None, description="对比产品")
    comparison_results: Optional[str] = Field(default=None, description="对比结果")


class PackagingDesignSpecificationParameters(BaseTemplateParameters):
    """包装设计仕样书参数"""
    product_name: Optional[str] = Field(default=None, description="产品名称")
    packaging_requirements: Optional[str] = Field(default=None, description="包装要求")


class UserManualSpecificationParameters(BaseTemplateParameters):
    """使用说明书仕样书参数"""
    product_name: Optional[str] = Field(default=None, description="产品名称")
    manual_requirements: Optional[str] = Field(default=None, description="说明书要求")


class ProjectPlanParameters(BaseTemplateParameters):
    """项目计划书参数"""
    project_scope: Optional[str] = Field(default=None, description="项目范围")
    project_timeline: Optional[str] = Field(default=None, description="项目时间线")

