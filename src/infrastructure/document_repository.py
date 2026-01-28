from typing import Optional, List


class LocalDocumentRepository:
    """兼容性占位类：本地文档仓储占位符（已被移除的实现的替代，用于避免导入错误）。"""

    def __init__(self, *args, **kwargs):
        pass

    def save(self, document) -> None:
        raise NotImplementedError("LocalDocumentRepository is a placeholder")

    def get_by_id(self, document_id: str):
        return None

    def list_by_project(self, project_id) -> List:
        return []


