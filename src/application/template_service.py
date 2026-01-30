"""Application layer use-cases for template operations."""
from typing import Dict, Optional, Any

from src.infrastructure.services_registry import template_service


def get_supported_templates(language: Optional[str] = None) -> Dict[str, Any]:
    """Return supported templates (use-case wrapper)."""
    return template_service.get_supported_templates(language)


def get_template_info(template_name: str, language: Optional[str] = None) -> Optional[Dict[str, object]]:
    """Return template info dict or None (use-case wrapper)."""
    return template_service.get_template_info(template_name, language)


