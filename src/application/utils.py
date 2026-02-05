from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional
import re

from src.infrastructure.services_registry import template_service
from src.config import settings


def generate_output_filename(template_name: str, parameters: Dict[str, Any], language: Optional[str] = None) -> str:
    """
    Generate output filename: <项目编号>-<模版文件名>-AI_<版本号>-<日期时间>.<文档后缀>
    
    格式: <project_id>-<template_display_name>-AI_<version>-<datetime>.<ext>
    """
    template_info = template_service.get_template_info(template_name, language)
    if not template_info:
        raise ValueError(f"模板 '{template_name}' 不存在")

    # 确定文件扩展名
    available_formats = template_info.get("available_formats", [])
    if "xlsx" in available_formats:
        extension = "xlsx"
    elif "docx" in available_formats:
        extension = "docx"
    else:
        extension = "xlsx"

    # 获取项目编号
    project_id = parameters.get("project_id") or parameters.get("project_number")
    if project_id:
        # 清理项目编号，只保留允许的字符
        clean_project_id = re.sub(r'[^\w\-\.\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]', '-', str(project_id))
    else:
        clean_project_id = "UNKNOWN_PROJECT"

    # 获取模板显示名称
    template_display_name = template_info.get("display_name", template_name)
    # 清理模板名称，只保留允许的字符
    clean_template_name = re.sub(r'[^\w\-\.\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]', '-', str(template_display_name))

    # 获取版本号
    version = parameters.get("version") or parameters.get("ver")
    if version:
        clean_version = re.sub(r'[^\w\-\.]', '-', str(version))
    else:
        clean_version = "v1.0"

    # 获取日期时间
    date_value = parameters.get("date") or parameters.get("created_date")
    if date_value:
        try:
            if isinstance(date_value, str):
                date_formats = [
                    "%Y-%m-%d %H:%M:%S",
                    "%Y-%m-%d %H:%M",
                    "%Y-%m-%d",
                    "%Y/%m/%d",
                    "%Y%m%d",
                    "%d/%m/%Y",
                    "%m/%d/%Y",
                ]
                parsed_date = None
                for fmt in date_formats:
                    try:
                        parsed_date = datetime.strptime(date_value, fmt)
                        break
                    except ValueError:
                        continue
                if parsed_date:
                    date_time_str = parsed_date.strftime("%Y%m%d-%H%M%S")
                else:
                    date_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")
            else:
                date_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")
        except Exception:
            date_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    else:
        date_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")

    # 组装文件名: <项目编号>-<模版文件名>-AI_<版本号>-<日期时间>
    filename = f"{clean_project_id}-{clean_template_name}-AI_{clean_version}-{date_time_str}"
    
    # 检查文件名长度限制
    max_length = settings.filename_max_length
    if len(filename) > max_length:
        # 如果超长，优先缩短模板名称
        max_template_length = max_length - len(clean_project_id) - len(clean_version) - len(date_time_str) - 10  # 10是分隔符和AI_的长度
        if max_template_length > 0 and len(clean_template_name) > max_template_length:
            clean_template_name = clean_template_name[:max_template_length]
            filename = f"{clean_project_id}-{clean_template_name}-AI_{clean_version}-{date_time_str}"
        
        # 如果还是超长，进一步缩短
        if len(filename) > max_length:
            # 保留最核心的部分：项目编号、版本号、日期时间
            filename = f"{clean_project_id[:20]}-AI_{clean_version}-{date_time_str}"

    return f"{filename}.{extension}"


def extract_project_info_from_filename(filename: str) -> Dict[str, str]:
    """
    Extract project_id and version from filename formatted as:
    project_version_template_datetime.ext
    """
    try:
        name_without_ext = filename.rsplit('.', 1)[0]
        parts = name_without_ext.split('_')
        if len(parts) >= 2:
            project_id = parts[0]
            version = parts[1]
            return {"project_id": project_id, "version": version}
        else:
            return {"project_id": "default", "version": "default"}
    except Exception:
        return {"project_id": "default", "version": "default"}


