from datetime import datetime
from pathlib import Path
from typing import Any, Dict
import re

from src.infrastructure.services_registry import template_service
from src.config import settings


def generate_output_filename(template_name: str, parameters: Dict[str, Any]) -> str:
    """
    Generate output filename: project_version_template_datetime.ext
    """
    template_info = template_service.get_template_info(template_name)
    if not template_info:
        raise ValueError(f"模板 '{template_name}' 不存在")

    available_formats = template_info.get("available_formats", [])
    if "xlsx" in available_formats:
        extension = "xlsx"
    elif "docx" in available_formats:
        extension = "docx"
    else:
        extension = "xlsx"

    filename_parts = []

    project_name = parameters.get("project_name") or parameters.get("project_id")
    if project_name:
        clean_project = re.sub(r'[^\w\-_\.\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]', '_', str(project_name))
        filename_parts.append(clean_project)
    else:
        filename_parts.append("UNKNOWN_PROJECT")

    version = parameters.get("version") or parameters.get("ver")
    if version:
        clean_version = re.sub(r'[^\w\-_\.]', '_', str(version))
        filename_parts.append(clean_version)
    else:
        filename_parts.append("v1.0")

    template_display_name = template_info.get("display_name", template_name)
    clean_template_name = re.sub(r'[^\w\-_\.\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]', '_', str(template_display_name))
    filename_parts.append(clean_template_name)

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
                    date_time_str = parsed_date.strftime("%Y%m%d_%H%M%S")
                    filename_parts.append(date_time_str)
                else:
                    filename_parts.append(datetime.now().strftime("%Y%m%d_%H%M%S"))
            else:
                filename_parts.append(datetime.now().strftime("%Y%m%d_%H%M%S"))
        except Exception:
            filename_parts.append(datetime.now().strftime("%Y%m%d_%H%M%S"))
    else:
        filename_parts.append(datetime.now().strftime("%Y%m%d_%H%M%S"))

    filename = "_".join(filename_parts)
    max_length = settings.filename_max_length
    if len(filename) > max_length:
        if len(filename_parts) >= 4:
            template_part = filename_parts[2]
            if len(template_part) > 20:
                template_part = template_part[:20]
            filename_parts[2] = template_part
            filename = "_".join(filename_parts)
            if len(filename) > max_length:
                essential_parts = [filename_parts[0], filename_parts[1], filename_parts[3]]
                filename = "_".join(essential_parts)

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


