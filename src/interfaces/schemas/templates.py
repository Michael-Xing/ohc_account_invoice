from typing import Any, List, Literal

from pydantic import BaseModel, Field, model_validator
from pydantic_core import PydanticUndefined


class NullToDefaultModel(BaseModel):
    """
    If client explicitly passes JSON `null`, treat it as "use default" when a default exists.
    - Missing field: keep Pydantic's normal default behavior
    - Explicit null: replace with field default / default_factory (if any); otherwise keep None
    """

    @model_validator(mode="before")
    @classmethod
    def _null_to_default(cls, data: Any):
        if not isinstance(data, dict):
            return data

        out = dict(data)
        for name, f in cls.model_fields.items():
            if name not in out or out[name] is not None:
                continue

            if f.default_factory is not None:
                out[name] = f.default_factory()
            elif f.default is not PydanticUndefined and f.default is not None:
                out[name] = f.default

        return out


class BaseTemplateParameters(NullToDefaultModel):
    project_number: str = Field(default="", description="项目编号")
    version: str = Field(default="", description="版本号")
    date: str = Field(default="", description="日期")
    author: str = Field(default="OHC账票AI助手", description="作者")
    document_type: str = Field(default="", description="文档类型")
    department: str = Field(default="", description="部门")
    language: str = Field(default="", description="语言代码 (zh/ja/en)，如果不指定则使用默认模板")


class FileListItem(NullToDefaultModel):
    """文件列表项"""
    file_number: str = Field(default="", description="文件编号")
    short_name: str = Field(default="", description="文件名称")
    stage: str = Field(default="", description="阶段")
    version: str = Field(default="", description="阶段")


class DHFIndexParameters(BaseTemplateParameters):
    """DHF INDEX参数"""
    theme_no: str = Field(default="", description="项目NO，填充到C3单元格")
    theme_name: str = Field(default="", description="商品类别，填充到C4单元格")
    product_model: str = Field(default="", description="商品型号，根据'/'分割，填充到C5单元格")
    sales_name: str = Field(default="", description="贩卖名称，根据'/'分割，填充到C6单元格")
    stage: str = Field(default="", description="阶段，拼接到C7单元格内容的后面")
    product_name: str = Field(default="", description="商品名")
    file_list: List[FileListItem] = Field(default_factory=list, description="文件列表")


class PTFFileNumberMapItem(NullToDefaultModel):
    """PTF INDEX 文件编号与简称"""
    file_number: str = Field(default="", description="文件编号")
    short_name: str = Field(default="", description="简称")


class PTFIndexParameters(BaseTemplateParameters):
    """PTF INDEX参数"""
    target_area: str = Field(default="", description="贩卖国家，多个值用半角逗号连接")
    file_number_map: List[PTFFileNumberMapItem] = Field(
        default_factory=list,
        description="文件编号与简称列表，按顺序对应模板中的数据行",
    )



class IndividualTestParameters(BaseTemplateParameters):
    """个别试验要项书/结果书公共参数"""
    phase: Literal["WS", "ES", "PP"] = Field(default="ES", description="试验阶段")


class IndividualTestSpecParameters(IndividualTestParameters):
    """个别试验要项书参数"""
    test_name: str = Field(default="", description="试验名称/试验项目（模板 D3）")
    test_number: str = Field(default="", description="试验编号（模板 K3）")
    theme_no: str = Field(default="", description="主题No（模板 D5）")
    product_model: str = Field(default="", description="产品型号（模板 K5）")
    meas_temperature: str = Field(default="", description="测定温度（模板 L8）")
    meas_humidity: str = Field(default="", description="测定湿度（模板 N8）")
    test_purpose: str = Field(default="", description="试验目的（模板 C37）")
    test_conditions: str = Field(default="", description="试验条件（模板 C41标题,C43填充）")
    test_method: str = Field(default="", description="试验方法（模板 C51标题,C53填充）")
    others: str = Field(default="", description="其他记录事项（模板 C60标题,C62填充）")
    admission_decision_standard: str = Field(default="", description="合否判定基准（模板 C70标题,C72填充）")
    source: str = Field(default="", description="出处（模板 C79标题,C81-C84填充）")


class IndividualTestResultParameters(IndividualTestParameters):
    """个别试验结果书参数"""
    test_name: str = Field(default="", description="试验名称/试验项目（模板 D3）")
    test_number: str = Field(default="", description="试验编号（模板 K3）")
    theme_no: str = Field(default="", description="主题No（模板 D5）")
    product_model: str = Field(default="", description="产品型号（模板 K5）")
    meas_temperature: str = Field(default="", description="测定温度（模板 L8）")
    meas_humidity: str = Field(default="", description="测定湿度（模板 N8）")
    test_purpose: str = Field(default="", description="试验目的（模板 C37）")
    test_conditions: str = Field(default="", description="试验条件（模板 C41标题,C43填充）")
    test_method: str = Field(default="", description="试验方法（模板 C51标题,C53填充）")
    others: str = Field(default="", description="其他记录事项（模板 C60标题,C62填充）")
    admission_decision_standard: str = Field(default="", description="合否判定基准（模板 C70标题,C72填充）")
    source: str = Field(default="", description="出处（模板 C79标题,C81-C84填充）")


class VerificationTestItem(NullToDefaultModel):
    """验证计划/结果书-试验项"""
    test_number: str = Field(default="", description="试验编号")
    test_name: str = Field(default="", description="试验名称")
    requirement_and_standard: str = Field(default="", description="适用标准/试验标准")
    individual_test_spec_number: str = Field(default="", description="个别试验要项书编号")
    individual_test_result_number: str = Field(default="", description="个别试验结果书编号")


class BaseVerificationParameters(BaseTemplateParameters):
    """验证计划/结果书公共参数"""
    phase: Literal["WS", "ES", "PP"] = Field(default="ES", description="阶段")
    test_list: List[VerificationTestItem] = Field(
        default_factory=list,
        description="试验列表，按表格列向下填充（每项对应模板一行）",
    )


class VerificationPlanParameters(BaseVerificationParameters):
    """验证计划书参数"""
    theme_no: str = Field(default="", description="项目NO")
    theme_name: str = Field(default="", description="项目名称")
    sales_name: str = Field(default="", description="贩卖名称")
    product_model: str = Field(default="", description="商品型式名")
    condition: str = Field(default="", description="条件")
    doubt_condition: str = Field(default="", description="疑问条件")


class VerificationResultParameters(BaseVerificationParameters):
    """验证结果书参数"""
    theme_no: str = Field(default="", description="项目NO")
    theme_name: str = Field(default="", description="项目名称")
    sales_name: str = Field(default="", description="贩卖名称")
    product_model: str = Field(default="", description="商品型式名")
    condition: str = Field(default="", description="条件")
    doubt_condition: str = Field(default="", description="疑问条件")
    verification_plan: str = Field(default="", description="验证计划书编号")  # 验证结果书独有


class BasicSpecificationParameters(BaseTemplateParameters):
    """基本规格书参数"""

    # ========== 基本信息 ==========
    product_model: str = Field(default="", description="商品型号，使用'/'分割，多值用于生成机种表和其他位置")
    sales_name: str = Field(default="", description="贩卖名称，使用'/'分割，多值用于生成机种表和其他位置")
    theme_no: str = Field(default="", description="项目NO")
    production_area: str = Field(default="", description="生产地")

    # ========== 产品概述 ==========
    product_overview: str = Field(default="", description="产品设计规格书中的产品概述，支持复合Markdown，包含文本、表格、图片")

    # ========== 参考文件和适用范围 ==========
    reference_document: str = Field(default="", description="参考文件，支持复合Markdown，包含文本、表格、图片")
    scope: str = Field(default="", description="适用范围，支持复合Markdown，包含文本、表格、图片")

    # ========== 术语定义 ==========
    definition_term_table: str = Field(default="", description="定义术语，支持复合Markdown，包含文本、表格、图片")

    # ========== 使用目的与对象 ==========
    use_purpose: str = Field(default="", description="商品目的/使用范围，支持复合Markdown，包含文本、表格、图片")
    intended_patients: str = Field(default="", description="目标用户或患者，支持复合Markdown，包含文本、表格、图片")
    intended_user: str = Field(default="", description="使用者或操作者，支持复合Markdown，包含文本、表格、图片")
    environment: str = Field(default="", description="使用场所，支持复合Markdown，包含文本、表格、图片")

    # ========== 尺寸及重量 ==========
    dimensions: str = Field(default="", description="尺寸，支持复合Markdown，包含文本、表格、图片")
    weight: str = Field(default="", description="质量，支持复合Markdown，包含文本、表格、图片")

    # ========== 法律法规及标准 ==========
    regulations_and_standards: str = Field(default="", description="法律法规及标准，支持复合Markdown，包含文本、表格、图片")

    # ========== 附带品或附件 ==========
    accessories: str = Field(default="", description="附带品或附件，支持复合Markdown，包含文本、表格、图片")

    # ========== 机能及性能 ==========
    performance_table: str = Field(default="", description="机能及性能，支持复合Markdown，包含文本、表格、图片")

    # ========== 安全性和风险评估 ==========
    safety_protection_info: str = Field(default="", description="安全性和风险评估，支持复合Markdown，包含文本、表格、图片")

    # ========== 出厂设定 ==========
    various_settings: str = Field(default="", description="出厂设定，支持复合Markdown，包含文本、表格、图片")

    # ========== 维修和保养 & 废弃处理 ==========
    maintenance: str = Field(default="", description="维修和保养，支持复合Markdown，包含文本、表格、图片")
    disposal: str = Field(default="", description="废弃处理，支持复合Markdown，包含文本、表格、图片")

    # ========== 环境条件 ==========
    environmental_conditions: str = Field(default="", description="环境条件，支持复合Markdown，包含文本、表格、图片")

    # ========== 文件编号 ==========
    dimensional_number: str = Field(default="", description="外形尺寸图文件编号")
    dustomer_specific_numner: str = Field(default="", description="客先仕样图")
    packing_instructions_number: str = Field(default="", description="梱包说明图")
    igs_number: str = Field(default="", description="IGS")
    instruction_manual_number: str = Field(default="", description="使用说明书图纸")
    packing_specification_number: str = Field(default="", description="包装仕样图文件编号")


class FollowUpDRMinutesParameters(BaseTemplateParameters):
    """跟进DR会议记录参数"""
    meeting_date: str = Field(default="", description="会议日期")
    meeting_location: str = Field(default="", description="会议地点")


class LabelingSpecificationParameters(BaseTemplateParameters):
    """标签仕样书-仕样确认书参数"""
    theme_no: str = Field(default="", description="项目NO，填入D5单元格")
    theme_name: str = Field(default="", description="项目名称，填入K5单元格")
    product_model_name: str = Field(default="", description="商品型式名，填入D7单元格")
    representative_model: str = Field(default="", description="代表型号，填入G11单元格")
    product_model: str = Field(default="", description="商品型式，填充到E17单元格")
    product_name: str = Field(default="", description="商品名，拼接到I8单元格内容后面")
    sales_name: str = Field(default="", description="贩卖名称，填充到E19单元格")
    production_area: str = Field(default="", description="生产地，如果=OMD则填固定值到E22，否则空白")
    target_area: str = Field(default="", description="贩卖国家，多个值用半角逗号连接，包含OHC时填固定值到G19和G26")
    sales_channel: str = Field(default="", description="販売チャネル，填固定值到E26，贩卖渠道只有“医療機関”时→ 400-889-0089,多种贩卖渠道时→ 400-770-9988")
    stage: str = Field(default="", description="阶段")
    related_file_info: List[FileListItem] = Field(default_factory=list, description="关联文件列表")
    address: str = Field(default="", description="生产地址，填入G17单元格")
    country: str = Field(default="", description="生产国家，填入G18单元格")
    phone: str = Field(default="", description="联系电话")

class ProductEnvironmentAssessmentParameters(BaseTemplateParameters):
    """产品环境评估要项书/结果书参数"""
    theme_no: str = Field(default="", description="项目NO，拼接到B5单元格内容后面")
    theme_name: str = Field(default="", description="商品类别（已废弃，不再使用）")
    product_model: str = Field(default="", description="商品型号，根据'/'分割，填充到22行的D～H列合并单元格")
    product_model_name: str = Field(default="", description="商品型号名，拼接到B7单元格内容后面")
    product_name: str = Field(default="", description="商品名，拼接到I5单元格内容后面")
    production_area: str = Field(default="", description="生产地，拼接到I7单元格内容后面")
    sales_name: str = Field(default="", description="贩卖名称，根据'/'分割，填充到22行的I～J列合并单元格")
    target_area: str = Field(default="", description="贩卖国家，填充到22行的K～M列合并单元格")
    remarks: str = Field(default="", description="备注（可选）")
    eta_schedule: str = Field(default="", description="ETA预定日志（可选）")


class ExistingProductComparisonParameters(BaseTemplateParameters):
    """与现有产品对比表参数"""
    comparison_products: str = Field(default="", description="对比产品")
    comparison_results: str = Field(default="", description="对比结果")


class PackagingDesignSpecificationParameters(BaseTemplateParameters):
    """包装设计仕样书参数"""
    theme_no: str = Field(default="", description="项目NO，填入C21单元格")
    theme_name: str = Field(default="", description="项目名称，填入E21单元格")
    product_model_name: str = Field(default="", description="商品型式名，填入L21单元格")
    sales_name: str = Field(default="", description="贩卖名称，填入C23单元格")
    related_file_info: List[FileListItem] = Field(
        default_factory=list,
        description="关联文件列表",
    )

class UserManualSpecificationParameters(BaseTemplateParameters):
    """使用说明书仕样书参数"""
    theme_no: str = Field(default="", description="项目NO，填入B19单元格")
    theme_name: str = Field(default="", description="项目名称，填入D19单元格")
    product_model_name: str = Field(default="", description="商品型式名，填入J19单元格")
    sales_name: str = Field(default="", description="贩卖名称，填入B21单元格")
    related_file_info: List[FileListItem] = Field(
        default_factory=list,
        description="关联文件列表",
    )


class ProjectPlanParameters(BaseTemplateParameters):
    """项目计划书参数"""
    theme_no: str = Field(default="", description="项目NO")
    theme_name: str = Field(default="", description="项目名称")
    product_model_name: str = Field(default="", description="商品型式名")
    sales_name: str = Field(default="", description="贩卖名称")
    target: str = Field(default="", description="商品目标")
    differentiation: str = Field(default="", description="差异化")
    # design: str = Field(default="", description="设计一览")
    # product_requirement: str = Field(default="", description="产品要件书")
    issue_result: str = Field(default="", description="课题分析·对策结果书")
    document_drawing_no: str = Field(default="", description="文件·图纸编号")
    document_drawing_rev: str = Field(default="", description="文件·图纸版本")
    software: str = Field(default="", description="软件开发计划书")
    function_module: str = Field(default="")
    p_plan: str = Field(default="")
    equipment_plan: str = Field(default="")
    engineering_plan: str = Field(default="")
    customer_service: str = Field(default="")
    approval_plan: str = Field(default="")
    risk_management: str = Field(default="")
    es_verification_plan: str = Field(default="")
    es_verification_result: str = Field(default="")
    igs: str = Field(default="")
    qc_engineering: str = Field(default="")
    new_product_confirmation: str = Field(default="", description="新商品开发确认书")
    requirement_spec: str = Field(default="", description="要求仕样书")
    product_design_spec: str = Field(default="", description="产品设计仕样书")
    schedule: str = Field(default="", description="日程表")
    doc_record_list: str = Field(default="", description="作成文件·记录一览表")
    target_fc: str = Field(default="", description="开发目标FC")
    color_instruction_record: str = Field(default="", description="颜色指示相关记录")