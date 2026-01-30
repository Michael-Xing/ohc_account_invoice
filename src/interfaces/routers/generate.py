from fastapi import APIRouter
from typing import Any

from src.application.generate_service import generate_document_internal
from src.interfaces.schemas import (
    DHFIndexParameters, PTFIndexParameters, ESIndividualTestSpecParameters,
    ESIndividualTestResultParameters, PPIndividualTestResultParameters,
    ESVerificationPlanParameters, ESVerificationResultParameters,
    PPVerificationPlanParameters, PPVerificationResultParameters,
    BasicSpecificationParameters, PPIndividualTestSpecParameters,
    FollowUpDRMinutesParameters, LabelingSpecificationParameters,
    ProductEnvironmentAssessmentParameters, ExistingProductComparisonParameters,
    PackagingDesignSpecificationParameters, UserManualSpecificationParameters,
    ProjectPlanParameters
)
from src.interfaces.schemas import GenerateDocumentResponse, GenerateDocumentRequest

router = APIRouter(prefix="", tags=["generate"])


@router.post("/generate", response_model=GenerateDocumentResponse, summary="生成文档", description="生成账票文档（通用接口）")
async def generate_document(request: GenerateDocumentRequest):
    return GenerateDocumentResponse(**generate_document_internal(request.template_name, request.parameters, request.language))


@router.post("/generate/dhf-index", response_model=GenerateDocumentResponse, summary="生成DHF INDEX", description="生成制作文档・图纸一览")
async def generate_dhf_index(parameters: DHFIndexParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("DHF_INDEX", params_dict, language))


@router.post("/generate/ptf-index", response_model=GenerateDocumentResponse, summary="生成PTF INDEX", description="生成PTF INDEX")
async def generate_ptf_index(parameters: PTFIndexParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PTF_INDEX", params_dict, language))


@router.post("/generate/es-individual-test-spec", response_model=GenerateDocumentResponse, summary="生成ES个别试验要项书", description="生成ES个别试验要项书")
async def generate_es_individual_test_spec(parameters: ESIndividualTestSpecParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("ES_INDIVIDUAL_TEST_SPEC", params_dict, language))


@router.post("/generate/es-individual-test-result", response_model=GenerateDocumentResponse, summary="生成ES个别试验结果书", description="生成ES个别试验结果书")
async def generate_es_individual_test_result(parameters: ESIndividualTestResultParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("ES_INDIVIDUAL_TEST_RESULT", params_dict, language))


@router.post("/generate/pp-individual-test-result", response_model=GenerateDocumentResponse, summary="生成PP个别试验结果书", description="生成PP个别试验结果书")
async def generate_pp_individual_test_result(parameters: PPIndividualTestResultParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PP_INDIVIDUAL_TEST_RESULT", params_dict, language))


@router.post("/generate/es-verification-plan", response_model=GenerateDocumentResponse, summary="生成ES验证计划书", description="生成ES验证计划书")
async def generate_es_verification_plan(parameters: ESVerificationPlanParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("ES_VERIFICATION_PLAN", params_dict, language))


@router.post("/generate/es-verification-result", response_model=GenerateDocumentResponse, summary="生成ES验证结果书", description="生成ES验证结果书")
async def generate_es_verification_result(parameters: ESVerificationResultParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("ES_VERIFICATION_RESULT", params_dict, language))


@router.post("/generate/pp-verification-plan", response_model=GenerateDocumentResponse, summary="生成PP验证计划书", description="生成PP验证计划书")
async def generate_pp_verification_plan(parameters: PPVerificationPlanParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PP_VERIFICATION_PLAN", params_dict, language))


@router.post("/generate/pp-verification-result", response_model=GenerateDocumentResponse, summary="生成PP验证结果书", description="生成PP验证结果书")
async def generate_pp_verification_result(parameters: PPVerificationResultParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PP_VERIFICATION_RESULT", params_dict, language))


@router.post("/generate/basic-specification", response_model=GenerateDocumentResponse, summary="生成基本规格书", description="生成基本规格书")
async def generate_basic_specification(parameters: BasicSpecificationParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("BASIC_SPECIFICATION", params_dict, language))


@router.post("/generate/pp-individual-test-spec", response_model=GenerateDocumentResponse, summary="生成PP个别试验要项书", description="生成PP个别试验要项书")
async def generate_pp_individual_test_spec(parameters: PPIndividualTestSpecParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PP_INDIVIDUAL_TEST_SPEC", params_dict, language))


@router.post("/generate/follow-up-dr-minutes", response_model=GenerateDocumentResponse, summary="生成跟进DR会议记录", description="生成跟进DR会议记录")
async def generate_follow_up_dr_minutes(parameters: FollowUpDRMinutesParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("FOLLOW_UP_DR_MINUTES", params_dict, language))


@router.post("/generate/labeling-specification", response_model=GenerateDocumentResponse, summary="生成标签规格书", description="生成标签规格书")
async def generate_labeling_specification(parameters: LabelingSpecificationParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("LABELING_SPECIFICATION", params_dict, language))


@router.post("/generate/product-environment-assessment", response_model=GenerateDocumentResponse, summary="生成产品环境评估要项书/结果书", description="生成产品环境评估要项书/结果书")
async def generate_product_environment_assessment(parameters: ProductEnvironmentAssessmentParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PRODUCT_ENVIRONMENT_ASSESSMENT", params_dict, language))


@router.post("/generate/existing-product-comparison", response_model=GenerateDocumentResponse, summary="生成与现有产品对比表", description="生成与现有产品对比表")
async def generate_existing_product_comparison(parameters: ExistingProductComparisonParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("EXISTING_PRODUCT_COMPARISON", params_dict, language))


@router.post("/generate/packaging-design-specification", response_model=GenerateDocumentResponse, summary="生成包装设计仕样书", description="生成包装设计仕样书")
async def generate_packaging_design_specification(parameters: PackagingDesignSpecificationParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PACKAGING_DESIGN_SPECIFICATION", params_dict, language))


@router.post("/generate/user-manual-specification", response_model=GenerateDocumentResponse, summary="生成使用说明书仕样书", description="生成使用说明书仕样书")
async def generate_user_manual_specification(parameters: UserManualSpecificationParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("USER_MANUAL_SPECIFICATION", params_dict, language))


@router.post("/generate/project-plan", response_model=GenerateDocumentResponse, summary="生成项目计划书", description="生成项目计划书")
async def generate_project_plan(parameters: ProjectPlanParameters):
    params_dict = parameters.model_dump()
    language = params_dict.pop("language", None)
    return GenerateDocumentResponse(**generate_document_internal("PROJECT_PLAN", params_dict, language))

