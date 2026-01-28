import os
import tempfile
from pathlib import Path

import pytest

# Ensure infrastructure init is skipped in test environment to avoid external calls
os.environ.setdefault("SKIP_INFRA_INIT", "1")
from src.application import generate_service as gs


def test_generate_document_success_remote(monkeypatch, tmp_path):
    # Mock template_service to validate name and write output file
    def validate(name):
        return True

    def gen_doc(template_name, parameters, output_path):
        output_path.write_bytes(b"ok")
        return True

    class DummyTemplateSvc:
        validate_template_name = staticmethod(validate)
        generate_document = staticmethod(gen_doc)

    class DummyStorage:
        def save_file(self, file_path, file_name, project_id=None, version=None):
            return True, "http://example.com/file", "ok"

    monkeypatch.setattr(gs, "template_service", DummyTemplateSvc())
    monkeypatch.setattr(gs, "storage_service", DummyStorage())
    monkeypatch.setattr(gs, "settings", type("S", (), {"storage_type": "minio"}))

    res = gs.generate_document_internal("DHF_INDEX", {"project_name": "P", "version": "v1"})
    assert res["success"] is True
    assert res["file_url"] == "http://example.com/file"


def test_generate_document_success_local(monkeypatch, tmp_path):
    # Mock template_service and settings for local storage branch
    def validate(name):
        return True

    def gen_doc(template_name, parameters, output_path):
        output_path.write_bytes(b"ok")
        return True

    class DummyTemplateSvc:
        validate_template_name = staticmethod(validate)
        generate_document = staticmethod(gen_doc)

    monkeypatch.setattr(gs, "template_service", DummyTemplateSvc())
    # settings with get_local_storage_path
    class S:
        storage_type = "local"

        @staticmethod
        def get_local_storage_path():
            p = Path(tempfile.mkdtemp())
            return p

    monkeypatch.setattr(gs, "settings", S())
    # storage_service not used in local branch
    monkeypatch.setattr(gs, "storage_service", None)

    res = gs.generate_document_internal("DHF_INDEX", {"project_name": "P", "version": "v1"})
    assert res["success"] is True
    assert res["file_content"] is not None


