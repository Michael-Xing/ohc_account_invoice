import os
os.environ.setdefault("SKIP_INFRA_INIT", "1")
from src.application import template_service as app_ts


def test_get_supported_templates():
    d = app_ts.get_supported_templates()
    assert isinstance(d, dict)
    assert "DHF_INDEX" in d


def test_get_template_info():
    info = app_ts.get_template_info("DHF_INDEX")
    assert info is not None
    assert info.get("name") == "DHF_INDEX"

