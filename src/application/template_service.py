"""Application layer use-cases for template operations."""
from typing import Dict, Optional

from src.infrastructure.services_registry import template_service


def get_supported_templates() -> Dict[str, str]:
    """Return supported templates (use-case wrapper)."""
    return template_service.get_supported_templates()


def get_template_info(template_name: str) -> Optional[Dict[str, object]]:
    """Return template info dict or None (use-case wrapper)."""
    return template_service.get_template_info(template_name)


