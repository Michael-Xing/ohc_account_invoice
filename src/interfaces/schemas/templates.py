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


class PTFIndexParameters(BaseTemplateParameters):
    """PTF INDEX参数"""
    target_area: str = Field(default="", description="贩卖国家，多个值用半角逗号连接")
    file_numbers: List[str] = Field(default_factory=list, description="文件编号列表，按顺序依次填充到模板对应行")



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
    test_conditions: str = Field(default="", description="试验条件（模板 C44）")


class IndividualTestResultParameters(IndividualTestParameters):
    """个别试验结果书参数"""
    test_name: str = Field(default="", description="试验名称/试验项目（模板 D3）")
    test_number: str = Field(default="", description="试验编号（模板 K3）")
    theme_no: str = Field(default="", description="主题No（模板 D5）")
    product_model: str = Field(default="", description="产品型号（模板 K5）")
    meas_temperature: str = Field(default="", description="测定温度（模板 L8）")
    meas_humidity: str = Field(default="", description="测定湿度（模板 N8）")
    test_purpose: str = Field(default="", description="试验目的（模板 C37）")
    test_conditions: str = Field(default="", description="试验条件（模板 C44）")


class BaseVerificationParameters(BaseTemplateParameters):
    """验证计划/结果书公共参数"""
    phase: Literal["WS", "ES", "PP"] = Field(default="ES", description="阶段")


class VerificationPlanParameters(BaseVerificationParameters):
    """验证计划书参数"""
    theme_no: str = Field(default="", description="项目NO")
    theme_name: str = Field(default="", description="项目名称")
    sales_name: str = Field(default="", description="贩卖名称")
    product_model_name: str = Field(default="", description="商品型式名")
    environment_temperature: str = Field(default="", description="环境温度")
    relative_humidity: str = Field(default="", description="相对湿度")
    test_voltage: str = Field(default="", description="试验电压")
    test_numbers: List[str] = Field(default_factory=list, description="试验编号列表，按表格列向下填充")
    test_names: List[str] = Field(default_factory=list, description="试验名称列表，按表格列向下填充")
    requirements_and_standards: List[str] = Field(default_factory=list, description="适用标准/试验标准列表，按表格列向下填充")


class VerificationResultParameters(BaseVerificationParameters):
    """验证结果书参数"""
    theme_no: str = Field(default="", description="项目NO")
    theme_name: str = Field(default="", description="项目名称")
    sales_name: str = Field(default="", description="贩卖名称")
    product_model_name: str = Field(default="", description="商品型式名")
    environment_temperature: str = Field(default="", description="环境温度")
    relative_humidity: str = Field(default="", description="相对湿度")
    test_voltage: str = Field(default="", description="试验电压")
    test_numbers: List[str] = Field(default_factory=list, description="试验编号列表，按表格列向下填充")
    test_names: List[str] = Field(default_factory=list, description="试验名称列表，按表格列向下填充")
    requirements_and_standards: List[str] = Field(default_factory=list, description="适用标准/试验标准列表，按表格列向下填充")


class BasicSpecificationServiceEnvironmentConditions(NullToDefaultModel):
    """基本规格书-使用环境及条件"""
    power_supply: str = Field(default="", description="电源")  # 使用环境及条件-电源
    use_temperature_humidity_range: str = Field(default="", description="使用湿度范围")  # 使用环境及条件-使用温湿度范围
    storage_and_transport_conditions: str = Field(default="", description="存储和运输环境")  # 使用环境及条件-存储和运输环境
    durability: str = Field(default="", description="耐久性")  # 使用环境及条件-耐久性


class BasicSpecificationSafetyProtectionInfo(NullToDefaultModel):
    """基本规格书-安全保护信息"""
    definitions_of_basic_safety: str = Field(default="", description="安全本质性能的定义(IEC60601-1中的基本性能)")  # 安全本质性能定义
    device_classification: str = Field(default="", description="安全设备分类")  # 安全设备分类
    equipment_safety_protection_and_warnings: str = Field(default="", description="对机器安全的保护和警告简述")  # 设备安全保护与警告
    safety_protection: str = Field(default="", description="安全防护")  # 安全防护
    safety_warning: str = Field(default="", description="安全警告")  # 安全警告
    biological_alarms: str = Field(default="", description="生理学警报")  # 生理学警报
    technical_alarms: str = Field(default="", description="技术警报")  # 技术警报


class BasicSpecificationVariousSettings(NullToDefaultModel):
    """基本规格书-各种设置"""
    default_equipment_setting: str = Field(default="", description="默认出场设置")  # 设备默认出厂设置
    date_time_settings: str = Field(default="", description="时间和日期设置")  # 时间日期设置


class BasicSpecificationMaintenanceAndDisposal(NullToDefaultModel):
    """基本规格书-保存维护和废弃处理"""
    maintenance: str = Field(default="", description="维护保存")  # 日常维护与保存
    disposal: str = Field(default="", description="废弃处理")  # 废弃处理


class BasicSpecificationParameters(BaseTemplateParameters):
    """基本规格书参数"""
    # 基本信息类字段
    product_model: str = Field(default="", description="商品型号，使用'/'分割，多值用于生成机种表和其他位置")  # 商品型号
    sales_name: str = Field(default="", description="贩卖名称，使用'/'分割，多值用于生成机种表和其他位置")  # 贩卖名称
    theme_no: str = Field(default="", description="项目NO")  # 项目编号
    production_area: str = Field(default="", description="生产地")  # 生产地

    # 参考文件和适用范围
    reference_document: str = Field(default="", description="参考文件")  # 参考文件
    scope: str = Field(default="", description="适用范围")  # 适用范围

    # 术语定义（Markdown表格）
    definition_term_table: str = Field(default="", description="术语定义，Markdown表格，需转换为Word表格")  # 术语定义表格

    # 使用目的与对象
    use_purpose: str = Field(default="", description="使用目的，支持Markdown语法")  # 使用目的
    intended_patients: str = Field(default="", description="对象患者，支持Markdown语法")  # 对象患者
    intended_user: str = Field(default="", description="使用对象，支持Markdown语法")  # 使用对象
    environment: str = Field(default="", description="适用环境，支持Markdown语法")  # 适用环境
    use_type: str = Field(default="", description="使用种类/医疗目的，支持Markdown语法")  # 使用种类/医疗目的

    # 机能构成（Markdown表格）
    component_table: str = Field(default="", description="机能构成，Markdown表格，需转换为Word表格")  # 机能构成表格

    # 外观图片列表（字符串形式，例如 "['image_url1', 'image_url2']"，在使用时会转换为列表）
    appearance_image: str = Field(default="", description="外观图URL列表字符串，将下载并按顺序插入到同一占位符位置")  # 外观图片URL列表

    # 其他规格信息
    dimensions_and_weight: str = Field(default="", description="尺寸及重量，支持Markdown语法")  # 尺寸及重量
    regulations_and_standards: str = Field(default="", description="政策法规，支持Markdown语法")  # 政策法规

    # 使用环境及条件
    service_environment_conditions: BasicSpecificationServiceEnvironmentConditions = Field(
        default_factory=BasicSpecificationServiceEnvironmentConditions,
        description="使用环境及条件",
    )

    # 材料与附属品
    main_unit: str = Field(default="", description="材料主题，支持Markdown语法")  # 设备本体相关说明
    accessories: str = Field(default="", description="附属品，支持Markdown语法")  # 附属品说明

    # 安全保护信息
    safety_protection_info: BasicSpecificationSafetyProtectionInfo = Field(
        default_factory=BasicSpecificationSafetyProtectionInfo,
        description="安全保护信息",
    )

    # 各种设置
    various_settings: BasicSpecificationVariousSettings = Field(
        default_factory=BasicSpecificationVariousSettings,
        description="各种设置",
    )

    # 标签与包装
    labeling: str = Field(default="", description="标签，支持Markdown语法")  # 标签信息
    packaging: str = Field(default="", description="打包说明，支持Markdown语法")  # 打包说明

    # 保存维护和废弃处理
    maintenance_and_disposal: BasicSpecificationMaintenanceAndDisposal = Field(
        default_factory=BasicSpecificationMaintenanceAndDisposal,
        description="保存维护和废弃处理",
    )

    # 功能说明（Markdown表格）
    function_table: str = Field(default="", description="功能说明，Markdown表格，需转换为Word表格")  # 功能说明表格

    # 功能块结构图（字符串形式，例如 "['image_url1', 'image_url2']"，在使用时会转换为列表）
    function_block_image: str = Field(default="", description="功能块图URL列表字符串，将下载并按顺序插入到同一占位符位置")  # 功能块图URL列表

    # 功能模块（Markdown表格，列项相同内容需要合并单元格）
    function_block_table: str = Field(default="", description="功能模块，Markdown表格，需转换为Word表格并按列合并相同内容")  # 功能模块表格

    # 性能说明（Markdown表格，列项相同内容需要合并单元格）
    performance_table: str = Field(default="", description="性能说明，Markdown表格，需转换为Word表格并按列合并相同内容")  # 性能说明表格


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
    design: str = Field(default="", description="设计一览")
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