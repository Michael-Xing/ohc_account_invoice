"""模板填充服务模块（已迁移到 infrastructure 层）"""

import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, Optional

from docx import Document
from openpyxl import load_workbook

from src.config import settings
from src.infrastructure.template_service import TemplateFillerStrategy, ExcelTemplateFiller, WordTemplateFiller

from src.domain.fillers.dhf_index_filler import DHFIndexFiller
from src.domain.fillers.product_environment_assessment_filler import ProductEnvironmentAssessmentFiller
from src.domain.fillers.basic_specification_filler import BasicSpecificationFiller
from src.domain.fillers.labeling_specification_filter import LabelingSpecificationFiller
from src.domain.fillers.packaging_design_specification_filler import PackagingDesignSpecificationFiller
from src.domain.fillers.user_manual_specification_filler import UserManualSpecificationFiller

# ... 其他填充器为简洁起见省略；原始实现已保留在代码库中 ...


class TemplateService:
    """模板填充服务"""
    # 模板配置：包含多语言显示名称和描述（display_names中的值即为模板文件名）
    SUPPORTED_TEMPLATES = {
        "DHF_INDEX": {
            "description": "文档和图纸一览表，用于管理项目相关的所有文档和图纸",
            "display_names": {
                "zh": "文件･图纸一览",
                "ja": "ドキュメント・図面一覧",
                "en": "Document and Drawing Index"
            }
        },
        "PTF_INDEX": {
            "description": "PTF索引表，用于管理产品测试文件",
            "display_names": {
                "zh": "PTF INDEX",
                "ja": "PTF INDEX",
                "en": "PTF Index"
            }
        },
        "ES_INDIVIDUAL_TEST_SPEC": {
            "description": "ES个别试验要项书，用于记录ES阶段的个别试验要求",
            "display_names": {
                "zh": "个别试验要项书",
                "ja": "ES個別試験要項書",
                "en": "ES Individual Test Specification"
            }
        },
        "ES_INDIVIDUAL_TEST_RESULT": {
            "description": "ES个别试验结果书，用于记录ES阶段的个别试验结果",
            "display_names": {
                "zh": "个别试验要项书",
                "ja": "ES個別試験結果書",
                "en": "ES Individual Test Result"
            }
        },
        "PP_INDIVIDUAL_TEST_RESULT": {
            "description": "PP个别试验结果书，用于记录PP阶段的个别试验结果",
            "display_names": {
                "zh": "个别试验要项书",
                "ja": "PP個別試験結果書",
                "en": "PP Individual Test Result"
            }
        },
        "PP_INDIVIDUAL_TEST_SPEC": {
            "description": "PP个别试验要项书，用于记录PP阶段的个别试验要求",
            "display_names": {
                "zh": "个别试验要项书",
                "ja": "PP個別試験要項書",
                "en": "PP Individual Test Specification"
            }
        },
        "ES_VERIFICATION_PLAN": {
            "description": "ES验证计划书，用于制定ES阶段的验证计划",
            "display_names": {
                "zh": "检证计划・结果书",
                "ja": "ES検証計画書",
                "en": "ES Verification Plan"
            }
        },
        "ES_VERIFICATION_RESULT": {
            "description": "ES验证结果书，用于记录ES阶段的验证结果",
            "display_names": {
                "zh": "检证计划・结果书",
                "ja": "ES検証結果書",
                "en": "ES Verification Result"
            }
        },
        "PP_VERIFICATION_PLAN": {
            "description": "PP验证计划书，用于制定PP阶段的验证计划",
            "display_names": {
                "zh": "检证计划・结果书",
                "ja": "PP検証計画書",
                "en": "PP Verification Plan"
            }
        },
        "PP_VERIFICATION_RESULT": {
            "description": "PP验证结果书，用于记录PP阶段的验证结果",
            "display_names": {
                "zh": "检证计划・结果书",
                "ja": "PP検証結果書",
                "en": "PP Verification Result"
            }
        },
        "BASIC_SPECIFICATION": {
            "description": "基本规格书，用于定义产品的基本规格和功能要求",
            "display_names": {
                "zh": "基本规格书",
                "ja": "基本仕様書",
                "en": "Basic Specification"
            }
        },
        # TODO找到模版和对应的文件
        "FOLLOW_UP_DR_MINUTES": {
            "description": "跟进DR会议记录，用于记录设计评审会议的跟进事项",
            "display_names": {
                "zh": "跟进DR会议记录",
                "ja": "フォローアップDR議事録",
                "en": "Follow-up DR Minutes"
            }
        },
        "LABELING_SPECIFICATION": {
            "description": "标签规格书，用于定义产品标签的规格和要求",
            "display_names": {
                "zh": "标签仕样书-仕样确认书",
                "ja": "ラベル仕様書",
                "en": "Labeling Specification"
            }
        },
        "PRODUCT_ENVIRONMENT_ASSESSMENT": {
            "description": "产品环境评估要项书/结果书，用于评估产品在不同环境下的适应性",
            "display_names": {
                "zh": "基本机种产品环境评估要項書-結果書",
                "ja": "基本機種製品環境アセスメント要項書／結果書",
                "en": "Product Environment Assessment"
            }
        },
        "PACKAGING_DESIGN_SPECIFICATION": {
            "description": "包装设计仕样书，用于定义产品包装的设计规格和要求",
            "display_names": {
                "zh": "包装设计仕样书",
                "ja": "包装設計仕様書",
                "en": "Packaging Design Specification"
            }
        },
        "USER_MANUAL_SPECIFICATION": {
            "description": "使用说明书仕样书，用于定义产品使用说明书的规格和要求",
            "display_names": {
                "zh": "使用说明书仕样书",
                "ja": "取扱説明書仕様書",
                "en": "User Manual Specification"
            }
        },
        "PROJECT_PLAN": {
            "description": "项目计划书，用于制定和管理项目的整体计划",
            "display_names": {
                "zh": "项目计划书",
                "ja": "プロジェクト計画書",
                "en": "Project Plan"
            }
        }
    }

    TEMPLATE_FILLER_MAPPING = {
        "DHF_INDEX": DHFIndexFiller(),
        "PRODUCT_ENVIRONMENT_ASSESSMENT": ProductEnvironmentAssessmentFiller(),
        "BASIC_SPECIFICATION": BasicSpecificationFiller(),
        "LABELING_SPECIFICATION": LabelingSpecificationFiller(),
        "PACKAGING_DESIGN_SPECIFICATION": PackagingDesignSpecificationFiller(),
        "USER_MANUAL_SPECIFICATION": UserManualSpecificationFiller(),
        # 其他模板使用默认策略
    }

    def __init__(self):
        self.template_base_path = Path(settings.template_base_path)

    def get_supported_templates(self, language: Optional[str] = None) -> Dict[str, Any]:
        """
        获取支持的模板列表
        
        Args:
            language: 语言代码 (zh/ja/en)，如果为None则返回包含描述信息的完整配置
            
        Returns:
            如果指定语言：返回 {template_name: display_name} 格式
            如果未指定语言：返回 {template_name: description} 格式
        """
        if language:
            # 返回指定语言的显示名称
            result = {}
            for template_name, config in self.SUPPORTED_TEMPLATES.items():
                display_name = config["display_names"].get(language, config["display_names"].get("zh", template_name))
                result[template_name] = display_name
            return result
        else:
            # 返回包含描述信息的配置
            result = {}
            for template_name, config in self.SUPPORTED_TEMPLATES.items():
                result[template_name] = config.get("description", "")
            return result

    def validate_template_name(self, template_name: str) -> bool:
        return template_name in self.SUPPORTED_TEMPLATES

    def get_template_path(self, template_name: str, file_type: Optional[str] = None, language: Optional[str] = None) -> Optional[Path]:
        """
        获取模板文件路径
        
        Args:
            template_name: 模板名称（模板标识符）
            file_type: 文件类型 (excel/word)，如果为None则自动检测
            language: 语言代码 (zh/ja/en)，如果为None则尝试所有语言
            
        Returns:
            模板文件路径，如果不存在则返回None
        """
        if not self.validate_template_name(template_name):
            return None
        
        template_config = self.SUPPORTED_TEMPLATES[template_name]
        display_names = template_config["display_names"]
        
        # 支持的语言列表
        supported_languages = ["zh", "ja", "en"]
        
        # 确定要查找的语言列表
        if language and language in supported_languages:
            languages_to_try = [language]
        else:
            languages_to_try = supported_languages
        
        # 确定要查找的文件类型列表
        if file_type:
            file_types_to_try = [file_type]
        else:
            # 自动检测：先尝试excel，再尝试word
            file_types_to_try = ["excel", "word"]
        
        # 按优先级查找：先按语言，再按文件类型
        for lang in languages_to_try:
            # 获取该语言的显示名称（即文件名，不含扩展名）
            file_name_base = display_names.get(lang)
            if not file_name_base:
                continue
            
            for ft in file_types_to_try:
                # 构建模板目录路径
                if Path(self.template_base_path).is_absolute():
                    template_dir = Path(self.template_base_path) / ft / lang
                else:
                    template_dir = Path.cwd() / self.template_base_path / ft / lang
                
                # 构建文件路径（使用display_name作为文件名）
                extension = "xlsx" if ft == "excel" else "docx" if ft == "word" else ft
                template_file = template_dir / f"{file_name_base}.{extension}"
                
                if template_file.exists():
                    return template_file
        
        return None

    def fill_excel_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        filler = ExcelTemplateFiller()
        return filler.fill_template(template_path, parameters, output_path)

    def fill_word_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        filler = WordTemplateFiller()
        return filler.fill_template(template_path, parameters, output_path)

    def generate_document(self, template_name: str, parameters: Dict[str, Any], output_path: Path, language: Optional[str] = None) -> bool:
        if not self.validate_template_name(template_name):
            return False
        filler = self.TEMPLATE_FILLER_MAPPING.get(template_name)
        if not filler:
            return self._generate_with_default_strategy(template_name, parameters, output_path, language)
        return self._generate_with_specific_strategy(filler, template_name, parameters, output_path, language)

    def _generate_with_specific_strategy(self, filler: TemplateFillerStrategy, template_name: str, 
                                       parameters: Dict[str, Any], output_path: Path, language: Optional[str] = None) -> bool:
        # 自动检测文件类型（先尝试excel，再尝试word）
        template_path = self.get_template_path(template_name, None, language)
        if template_path:
            return filler.fill_template(template_path, parameters, output_path)
        return False

    def _generate_with_default_strategy(self, template_name: str, parameters: Dict[str, Any], 
                                      output_path: Path, language: Optional[str] = None) -> bool:
        # 自动检测文件类型（先尝试excel，再尝试word）
        template_path = self.get_template_path(template_name, None, language)
        if not template_path:
            return False
        
        # 根据文件扩展名选择填充器
        if template_path.suffix == ".xlsx":
            return self.fill_excel_template(template_path, parameters, output_path)
        elif template_path.suffix == ".docx":
            return self.fill_word_template(template_path, parameters, output_path)
        return False

    def get_template_info(self, template_name: str, language: Optional[str] = None) -> Optional[Dict[str, Any]]:
        if not self.validate_template_name(template_name):
            return None
        
        template_config = self.SUPPORTED_TEMPLATES[template_name]
        
        # 根据语言获取显示名称，如果没有指定语言或语言不存在，使用中文作为默认
        if language and language in template_config["display_names"]:
            display_name = template_config["display_names"][language]
        else:
            display_name = template_config["display_names"].get("zh", template_name)
        
        info = {
            "name": template_name,
            "display_name": display_name,
            "description": template_config["description"],
            "available_formats": [],
            "filler_strategy": self._get_filler_strategy_name(template_name),
            "features": self._get_template_features(template_name),
            "available_languages": self._get_available_languages(template_name),
            "display_names": template_config["display_names"]
        }
        
        # 检查模板文件格式
        if language:
            # 如果指定了语言，检查该语言下的模板文件
            template_path = self.get_template_path(template_name, None, language)
            if template_path:
                if template_path.suffix == ".xlsx":
                    info["available_formats"].append("xlsx")
                elif template_path.suffix == ".docx":
                    info["available_formats"].append("docx")
        else:
            # 如果没有指定语言，检查所有语言下的模板文件
            for lang in ["zh", "ja", "en"]:
                template_path = self.get_template_path(template_name, None, lang)
                if template_path:
                    if template_path.suffix == ".xlsx" and "xlsx" not in info["available_formats"]:
                        info["available_formats"].append("xlsx")
                    elif template_path.suffix == ".docx" and "docx" not in info["available_formats"]:
                        info["available_formats"].append("docx")
        
        return info
    
    def _get_available_languages(self, template_name: str) -> list:
        """获取模板可用的语言列表"""
        available_languages = []
        supported_languages = ["zh", "ja", "en"]
        
        for lang in supported_languages:
            template_path = self.get_template_path(template_name, None, lang)
            if template_path:
                available_languages.append(lang)
        
        # 如果没有找到任何语言版本，检查根目录是否有模板（向后兼容）
        if not available_languages:
            template_path = self.get_template_path(template_name, None, None)
            if template_path:
                available_languages.append("default")
        
        return available_languages

    def _get_filler_strategy_name(self, template_name: str) -> str:
        filler = self.TEMPLATE_FILLER_MAPPING.get(template_name)
        if filler:
            return filler.__class__.__name__
        return "DefaultStrategy"

    def _get_template_features(self, template_name: str) -> list:
        features = ["basic_placeholder_replacement"]
        filler = self.TEMPLATE_FILLER_MAPPING.get(template_name)
        if filler:
            # simplified feature mapping
            pass
        return features


