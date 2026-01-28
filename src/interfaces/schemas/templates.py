from typing import Optional
from pydantic import BaseModel, Field


class BaseTemplateParameters(BaseModel):
    project_name: Optional[str] = Field(default=None, description="项目名称")
    version: Optional[str] = Field(default=None, description="版本号")
    date: Optional[str] = Field(default=None, description="日期")
    author: Optional[str] = Field(default="OHC账票AI助手", description="作者")
    document_type: Optional[str] = Field(default=None, description="文档类型")
    department: Optional[str] = Field(default=None, description="部门")


class DHFIndexParameters(BaseTemplateParameters):
    pass


class PTFIndexParameters(BaseTemplateParameters):
    pass


class ESIndividualTestSpecParameters(BaseTemplateParameters):
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
    product_name: Optional[str] = None


class ProductEnvironmentAssessmentParameters(BaseTemplateParameters):
    product_name: Optional[str] = None


class ExistingProductComparisonParameters(BaseTemplateParameters):
    comparison_products: Optional[str] = None
    comparison_results: Optional[str] = None


