class TemplateNotFoundError(Exception):
    """Raised when a requested template does not exist."""
    pass


class TemplateGenerationError(Exception):
    """Raised when template generation fails."""
    pass


class StorageError(Exception):
    """Raised when storage (upload/save) fails."""
    pass


