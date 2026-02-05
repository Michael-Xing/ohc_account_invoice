import os
import json

# make sure infra init is skipped during import time
os.environ["SKIP_INFRA_INIT"] = "1"

from fastapi.testclient import TestClient

from src.main import app
from src.infrastructure import services_registry
from src.application import generate_service as generate_service_module
from src.application import utils as utils_module


class DummyStorage:
    def save_file(self, local_path, filename, project_id=None, version=None):
        # mimic successful remote storage save
        return True, f"https://example.com/{filename}", "saved"


def test_generate_document_e2e(monkeypatch, tmp_path):
    # monkeypatch storage and template services
    monkeypatch.setattr(services_registry, "storage_service", DummyStorage())
    # also patch references held inside application module to ensure runtime uses our mocks
    monkeypatch.setattr(generate_service_module, "storage_service", DummyStorage())

    class DummyTemplateService:
        def get_supported_templates(self):
            return {"TEST": "Test Template"}

        def get_template_info(self, name, language=None):
            return {"name": "TEST", "display_name": "Test Template", "available_formats": ["xlsx"]}
        
        def validate_template_name(self, name):
            return name == "TEST"

        def generate_document(self, template_name, parameters, output_path, language=None):
            # write a small xlsx-like content (not a real xlsx) so storage/read can proceed
            output_path.write_bytes(b"PK\\x03\\x04test")
            return True

    monkeypatch.setattr(services_registry, "template_service", DummyTemplateService())
    monkeypatch.setattr(generate_service_module, "template_service", DummyTemplateService())
    monkeypatch.setattr(utils_module, "template_service", DummyTemplateService())

    client = TestClient(app)

    payload = {
        "template_name": "TEST",
        "parameters": {
            "project_number": "E2E Project",
            "version": "1.0",
            "date": "2026-01-01",
            "author": "E2E Tester",
            "product_name": "Product",
            "overview": "overview text",
            "technical_requirements": "req",
            "acceptance_criteria": "criteria"
        }
    }

    resp = client.post("/generate", json=payload)
    assert resp.status_code == 200
    data = resp.json()
    assert data.get("success") is True
    assert data.get("file_url") is not None


