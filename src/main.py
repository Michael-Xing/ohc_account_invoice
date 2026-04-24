"""OHC账票生成FastAPI服务主模块"""

import base64
import logging
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field, field_validator, ConfigDict

from src.config import settings
from src.interfaces.middleware.auth import AuthMiddleware, create_auth_middleware

# Use uvicorn's logger name so it follows uvicorn log configuration.
logger = logging.getLogger("uvicorn.error")

# 创建FastAPI应用
app = FastAPI(
    title=settings.app_name,
    version=settings.app_version,
    description=settings.app_description,
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json"
)

# Include routers and application/infrastructure modules (use absolute imports only)
from src.interfaces.routers.templates import router as templates_router
from src.infrastructure.services_registry import template_service, storage_service
from src.interfaces.schemas import (
    GenerateDocumentRequest, GenerateDocumentResponse,
    TemplateInfoResponse, ServiceConfigResponse, HealthCheckResponse
)
from src.application.utils import generate_output_filename
from src.interfaces.routers.generate import router as generate_router
from src.interfaces.routers.system import router as system_router

app.include_router(templates_router)
app.include_router(generate_router)
app.include_router(system_router)

# 注册认证中间件
auth_middleware_config = create_auth_middleware()
if auth_middleware_config:
    app.add_middleware(
        AuthMiddleware,
        sso_validator=auth_middleware_config.sso_validator,
        api_key_validator=auth_middleware_config.api_key_validator,
        skip_paths=auth_middleware_config.skip_paths,
    )
    logger.info("认证中间件已注册")
else:
    logger.info("未启用认证中间件")


@app.exception_handler(RequestValidationError)
async def request_validation_error_handler(request: Request, exc: RequestValidationError):
    """
    Log detailed request validation errors (422) to help debugging client payload issues.
    """
    # exc.errors() is a list of structured items: {loc, msg, type, ...}
    errors = exc.errors()
    # Put details into the message too, so it works with any log formatter.
    logger.warning(
        "Request validation failed: method=%s path=%s query=%s client=%s errors=%s",
        request.method,
        str(request.url.path),
        str(request.url.query),
        getattr(request.client, "host", None),
        errors,
        extra={
            "http": {
                "method": request.method,
                "path": str(request.url.path),
                "query": str(request.url.query),
                "client": getattr(request.client, "host", None),
            },
            "validation": {
                "errors": errors,
            },
        },
    )
    return JSONResponse(status_code=422, content={"detail": errors})


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    """
    全局异常处理器 - 捕获所有未处理的异常，记录完整堆栈信息。
    """
    import traceback
    exc_type = type(exc).__name__
    exc_msg = str(exc)
    exc_traceback = traceback.format_exc()
    
    # 记录完整的错误信息
    logger.error(
        "Unhandled exception: type=%s message=%s traceback=%s path=%s method=%s client=%s",
        exc_type,
        exc_msg,
        exc_traceback,
        request.url.path,
        request.method,
        getattr(request.client, "host", None),
        extra={
            "http": {
                "method": request.method,
                "path": str(request.url.path),
                "client": getattr(request.client, "host", None),
            },
            "exception": {
                "type": exc_type,
                "message": exc_msg,
                "traceback": exc_traceback,
            },
        },
    )
    
    # 返回 500 错误，并包含错误信息（调试模式下更详细）
    content = {
        "detail": f"服务器内部错误: {exc_msg}",
        "error_type": exc_type,
    }
    
    # 如果开启了调试模式，返回更详细的信息
    if settings.debug:
        content["traceback"] = exc_traceback
    
    return JSONResponse(status_code=500, content=content)

# Optional Sentry monitoring initialization (if configured and package available)
try:
    if getattr(settings, "sentry_dsn", None):
        try:
            import importlib

            sentry_sdk = importlib.import_module("sentry_sdk")
            starlette_mod = importlib.import_module("sentry_sdk.integrations.starlette")
            StarletteIntegration = getattr(starlette_mod, "StarletteIntegration")

            sentry_sdk.init(
                dsn=settings.sentry_dsn,
                environment=getattr(settings, "sentry_environment", None),
                integrations=[StarletteIntegration()],
            )
        except Exception:
            # missing package or initialization failed; skip monitoring initialization
            pass
except Exception:
    # avoid failing app import if monitoring setup errors
    pass


if __name__ == "__main__":
    import uvicorn
    # uvicorn.run() 不支持 debug 参数，使用 log_level 来控制日志级别
    log_level = "debug" if settings.debug else "info"
    uvicorn.run(
        "src.main:app",
        host=settings.host,
        port=settings.port,
        reload=settings.reload,
        log_level=log_level
    )