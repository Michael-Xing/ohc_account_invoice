from pathlib import Path
import tempfile
from datetime import datetime
from typing import Any, Dict, Optional

from src.infrastructure.services_registry import template_service, storage_service
from src.config import settings
from src.application.utils import generate_output_filename
from src.application.logging_config import get_logger
from src.application.errors import TemplateNotFoundError, TemplateGenerationError, StorageError


def generate_document_internal(template_name: str, parameters: Dict[str, Any], language: Optional[str] = None) -> Dict[str, Optional[Any]]:
    """核心生成逻辑（从 main 中提取）；返回适用于响应模型的字典。"""
    result = {
        "success": False,
        "message": "",
        "file_name": None,
        "file_url": None,
        "storage_type": None,
        "project_id": None,
        "version": None,
    }

    logger = get_logger("application.generate")
    logger.info("generate_document_internal start: template=%s, language=%s", template_name, language)
    try:
        # Validate template
        if not template_service.validate_template_name(template_name):
            raise TemplateNotFoundError(f"不支持的模板: {template_name}")

        # Generate output filename
        output_filename = generate_output_filename(template_name, parameters, language)
        logger.info("Generated filename %s", output_filename)

        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{output_filename.split('.')[-1]}") as temp_file:
            temp_path = Path(temp_file.name)

        # Generate document using template service
        success = template_service.generate_document(template_name, parameters, temp_path, language)
        if not success:
            raise TemplateGenerationError("文档生成失败，请检查模板和参数")

        # Extract project/version
        project_id = parameters.get("project_name") or parameters.get("project_id")
        version = parameters.get("version") or parameters.get("ver")

        # Store file
        # 处理枚举类型：如果是枚举，使用 .value 获取字符串值；否则直接转换为字符串
        storage_type_val = None
        if hasattr(settings, "storage_type") and settings.storage_type is not None:
            if hasattr(settings.storage_type, "value"):
                storage_type_val = str(settings.storage_type.value).lower()
            else:
                storage_type_val = str(settings.storage_type).lower()
        
        if storage_type_val == "local":
            # local storage: use settings to write file locally
            try:
                storage_path = settings.get_local_storage_path()
                # create folder structure if applicable
                target_dir = storage_path / (project_id or "default") / (version or "default")
                target_dir.mkdir(parents=True, exist_ok=True)
                target_path = target_dir / output_filename
                temp_path.replace(target_path)
                file_url = str(target_path)
                success = True
                message = "文件保存成功"
            except Exception as e:
                success = False
                file_url = None
                message = str(e)
        else:
            if storage_service is None:
                raise StorageError("storage_service is not initialized")
            success, file_url, message = storage_service.save_file(
                temp_path,
                output_filename,
                project_id=project_id,
                version=version
            )

        temp_path.unlink(missing_ok=True)

        if not success:
            raise StorageError(message or "文件存储失败")

        result.update({
            "success": True,
            "message": "文档生成成功",
            "file_name": output_filename,
            "file_url": file_url,
            "storage_type": settings.storage_type.value if hasattr(settings.storage_type, "value") else settings.storage_type,
            "project_id": project_id,
            "version": version,
        })
        logger.info("generate_document_internal success: %s", output_filename)
        return result

    except TemplateNotFoundError as e:
        logger.warning("Template not found: %s", e)
        result["success"] = False
        result["message"] = str(e)
        return result
    except TemplateGenerationError as e:
        logger.error("Template generation failed: %s", e)
        result["success"] = False
        result["message"] = str(e)
        return result
    except StorageError as e:
        logger.error("Storage failed: %s", e)
        result["success"] = False
        result["message"] = str(e)
        return result
    except Exception as e:
        logger.exception("Unexpected error in generate_document_internal")
        result["success"] = False
        result["message"] = f"文档生成失败: {str(e)}"
        return result


