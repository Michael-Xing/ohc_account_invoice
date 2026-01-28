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
    ProductEnvironmentAssessmentParameters, ExistingProductComparisonParameters
)
from src.interfaces.schemas import GenerateDocumentResponse, GenerateDocumentRequest

router = APIRouter(prefix="", tags=["generate"])


@router.post("/generate", response_model=GenerateDocumentResponse, summary="生成文档", description="生成账票文档（通用接口）")
async def generate_document(request: GenerateDocumentRequest):
    return GenerateDocumentResponse(**generate_document_internal(request.template_name, request.parameters))


@router.post("/generate/dhf-index", response_model=GenerateDocumentResponse, summary="生成DHF INDEX", description="生成制作文档・图纸一览")
async def generate_dhf_index(parameters: DHFIndexParameters):
    return GenerateDocumentResponse(**generate_document_internal("DHF_INDEX", parameters.model_dump()))


@router.post("/generate/ptf-index", response_model=GenerateDocumentResponse, summary="生成PTF INDEX", description="生成PTF INDEX")
async def generate_ptf_index(parameters: PTFIndexParameters):
    return GenerateDocumentResponse(**generate_document_internal("PTF_INDEX", parameters.model_dump()))


@router.post("/generate/es-individual-test-spec", response_model=GenerateDocumentResponse, summary="生成ES个别试验要项书", description="生成ES个别试验要项书")
async def generate_es_individual_test_spec(parameters: ESIndividualTestSpecParameters):
    return GenerateDocumentResponse(**generate_document_internal("ES_INDIVIDUAL_TEST_SPEC", parameters.model_dump()))


@router.post("/generate/es-individual-test-result", response_model=GenerateDocumentResponse, summary="生成ES个别试验结果书", description="生成ES个别试验结果书")
async def generate_es_individual_test_result(parameters: ESIndividualTestResultParameters):
    return GenerateDocumentResponse(**generate_document_internal("ES_INDIVIDUAL_TEST_RESULT", parameters.model_dump()))


@router.post("/generate/pp-individual-test-result", response_model=GenerateDocumentResponse, summary="生成PP个别试验结果书", description="生成PP个别试验结果书")
async def generate_pp_individual_test_result(parameters: PPIndividualTestResultParameters):
    return GenerateDocumentResponse(**generate_document_internal("PP_INDIVIDUAL_TEST_RESULT", parameters.model_dump()))


@router.post("/generate/es-verification-plan", response_model=GenerateDocumentResponse, summary="生成ES验证计划书", description="生成ES验证计划书")
async def generate_es_verification_plan(parameters: ESVerificationPlanParameters):
    return GenerateDocumentResponse(**generate_document_internal("ES_VERIFICATION_PLAN", parameters.model_dump()))


@router.post("/generate/es-verification-result", response_model=GenerateDocumentResponse, summary="生成ES验证结果书", description="生成ES验证结果书")
async def generate_es_verification_result(parameters: ESVerificationResultParameters):
    return GenerateDocumentResponse(**generate_document_internal("ES_VERIFICATION_RESULT", parameters.model_dump()))


@router.post("/generate/pp-verification-plan", response_model=GenerateDocumentResponse, summary="生成PP验证计划书", description="生成PP验证计划书")
async def generate_pp_verification_plan(parameters: PPVerificationPlanParameters):
    return GenerateDocumentResponse(**generate_document_internal("PP_VERIFICATION_PLAN", parameters.model_dump()))


@router.post("/generate/pp-verification-result", response_model=GenerateDocumentResponse, summary="生成PP验证结果书", description="生成PP验证结果书")
async def generate_pp_verification_result(parameters: PPVerificationResultParameters):
    return GenerateDocumentResponse(**generate_document_internal("PP_VERIFICATION_RESULT", parameters.model_dump()))


@router.post("/generate/basic-specification", response_model=GenerateDocumentResponse, summary="生成基本规格书", description="生成基本规格书")
async def generate_basic_specification(parameters: BasicSpecificationParameters):
    return GenerateDocumentResponse(**generate_document_internal("BASIC_SPECIFICATION", parameters.model_dump()))


@router.post("/generate/pp-individual-test-spec", response_model=GenerateDocumentResponse, summary="生成PP个别试验要项书", description="生成PP个别试验要项书")
async def generate_pp_individual_test_spec(parameters: PPIndividualTestSpecParameters):
    return GenerateDocumentResponse(**generate_document_internal("PP_INDIVIDUAL_TEST_SPEC", parameters.model_dump()))


@router.post("/generate/follow-up-dr-minutes", response_model=GenerateDocumentResponse, summary="生成跟进DR会议记录", description="生成跟进DR会议记录")
async def generate_follow_up_dr_minutes(parameters: FollowUpDRMinutesParameters):
    return GenerateDocumentResponse(**generate_document_internal("FOLLOW_UP_DR_MINUTES", parameters.model_dump()))


@router.post("/generate/labeling-specification", response_model=GenerateDocumentResponse, summary="生成标签规格书", description="生成标签规格书")
async def generate_labeling_specification(parameters: LabelingSpecificationParameters):
    return GenerateDocumentResponse(**generate_document_internal("LABELING_SPECIFICATION", parameters.model_dump()))


@router.post("/generate/product-environment-assessment", response_model=GenerateDocumentResponse, summary="生成产品环境评估要项书/结果书", description="生成产品环境评估要项书/结果书")
async def generate_product_environment_assessment(parameters: ProductEnvironmentAssessmentParameters):
    return GenerateDocumentResponse(**generate_document_internal("PRODUCT_ENVIRONMENT_ASSESSMENT", parameters.model_dump()))


@router.post("/generate/existing-product-comparison", response_model=GenerateDocumentResponse, summary="生成与现有产品对比表", description="生成与现有产品对比表")
async def generate_existing_product_comparison(parameters: ExistingProductComparisonParameters):
    return GenerateDocumentResponse(**generate_document_internal("EXISTING_PRODUCT_COMPARISON", parameters.model_dump()))


