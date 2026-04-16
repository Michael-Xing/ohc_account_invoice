"""认证中间件模块

提供灵活的认证机制，支持两种认证方式：
1. SSO服务器认证（Authorization头）
2. API Key认证（x-api-key头）

开发环境可跳过认证，生产环境需要强制认证。
"""

import asyncio
import json
import logging
import time
from datetime import datetime
from typing import Optional, Callable, Dict, Any, List
from dataclasses import dataclass

import httpx
from fastapi import Request, HTTPException
from fastapi.responses import JSONResponse
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.types import ASGIApp

from src.config import settings

logger = logging.getLogger("uvicorn.error")


@dataclass
class AuthResult:
    """认证结果"""
    is_authorized: bool
    user_info: Optional[Dict[str, Any]] = None
    auth_type: str = "none"
    error_message: Optional[str] = None


class SSOAuthValidator:
    """SSO认证验证器"""

    def __init__(self, verify_url: str, timeout: int = 10):
        self.verify_url = verify_url
        self.timeout = timeout
        self._cache: Dict[str, tuple[bool, Optional[Dict], float]] = {}
        self.cache_ttl = 300  # 缓存5分钟

    async def verify(self, authorization: str) -> AuthResult:
        """
        验证SSO令牌

        Args:
            authorization: Authorization头值（Bearer token）

        Returns:
            AuthResult: 认证结果
        """
        if not self.verify_url:
            logger.warning("Authorization验证URL未配置，跳过Authorization认证")
            return AuthResult(is_authorized=False, error_message="SSO服务未配置")

        # 检查缓存
        cache_key = authorization
        if cache_key in self._cache:
            cached_result, user_info, cached_time = self._cache[cache_key]
            if time.time() - cached_time < self.cache_ttl:
                return AuthResult(
                    is_authorized=cached_result,
                    user_info=user_info,
                    auth_type="sso"
                )

        try:
            # 构造SSO验证请求
            request_data = {
                "key": "OMRON",
                "value": "Password_T_T_Verify",
            }

            async with httpx.AsyncClient(timeout=self.timeout, verify=False) as client:
                response = await client.post(
                    self.verify_url,
                    json=request_data,
                    headers={
                        "Authorization": authorization,
                        "Content-Type": "application/json",
                    },
                )

                if response.status_code != 200:
                    logger.warning(f"SSO验证失败: HTTP {response.status_code}")
                    return AuthResult(
                        is_authorized=False,
                        error_message=f"SSO验证失败: HTTP {response.status_code}"
                    )

                # 解析响应
                result = response.json()
                if result.get("result") == "authorized":
                    user_info = {
                        "authenticated": True,
                        "source": "sso",
                        "timestamp": datetime.now().isoformat(),
                    }
                    # 缓存结果
                    self._cache[cache_key] = (True, user_info, time.time())
                    return AuthResult(
                        is_authorized=True,
                        user_info=user_info,
                        auth_type="sso"
                    )
                else:
                    return AuthResult(
                        is_authorized=False,
                        error_message="SSO验证未授权"
                    )

        except asyncio.TimeoutError:
            logger.error("SSO验证超时")
            return AuthResult(is_authorized=False, error_message="SSO验证超时")
        except Exception as e:
            logger.error(f"SSO验证异常: {e}")
            return AuthResult(is_authorized=False, error_message=f"SSO验证异常: {str(e)}")


class APIKeyValidator:
    """API Key验证器"""

    def __init__(self, valid_keys: List[str]):
        self.valid_keys = set(valid_keys)
        self._cache: Dict[str, tuple[bool, Optional[Dict], float]] = {}
        self.cache_ttl = 300  # 缓存5分钟

    async def verify(self, api_key: str) -> AuthResult:
        """
        验证API Key

        Args:
            api_key: API密钥

        Returns:
            AuthResult: 认证结果
        """
        if not api_key:
            return AuthResult(is_authorized=False, error_message="缺少API Key")

        # 检查缓存
        if api_key in self._cache:
            cached_result, user_info, cached_time = self._cache[api_key]
            if time.time() - cached_time < self.cache_ttl:
                return AuthResult(
                    is_authorized=cached_result,
                    user_info=user_info,
                    auth_type="api_key"
                )

        # 验证API Key
        if api_key in self.valid_keys:
            user_info = {
                "authenticated": True,
                "source": "api_key",
                "api_key_prefix": api_key[:8] + "...",
                "timestamp": datetime.now().isoformat(),
            }
            self._cache[api_key] = (True, user_info, time.time())
            return AuthResult(
                is_authorized=True,
                user_info=user_info,
                auth_type="api_key"
            )
        else:
            return AuthResult(
                is_authorized=False,
                error_message="无效的API Key"
            )


@dataclass
class AuthMiddlewareConfig:
    """认证中间件配置（用于传递给 add_middleware）"""
    sso_validator: Optional[SSOAuthValidator] = None
    api_key_validator: Optional[APIKeyValidator] = None
    skip_paths: List[str] = None


class AuthMiddleware(BaseHTTPMiddleware):
    """认证中间件"""

    def __init__(
        self,
        app: ASGIApp,
        sso_validator: Optional[SSOAuthValidator] = None,
        api_key_validator: Optional[APIKeyValidator] = None,
        skip_paths: Optional[List[str]] = None,
    ):
        super().__init__(app)
        self.sso_validator = sso_validator
        self.api_key_validator = api_key_validator
        self.skip_paths = skip_paths or []

        # 记录配置信息
        logger.info("认证中间件初始化")
        logger.info(f"  SSO认证: {'启用' if sso_validator else '禁用'}")
        logger.info(f"  API Key认证: {'启用' if api_key_validator else '禁用'}")
        logger.info(f"  跳过认证路径: {self.skip_paths}")

    def _should_skip_auth(self, path: str) -> bool:
        """检查路径是否跳过认证"""
        for skip_path in self.skip_paths:
            if path == skip_path or path.startswith(skip_path + "/"):
                return True
        return False

    async def _get_auth_headers(self, request: Request) -> tuple[Optional[str], Optional[str]]:
        """从请求头获取两种认证信息"""
        auth_header = request.headers.get("Authorization")
        api_key = request.headers.get("x-api-key")
        return auth_header, api_key

    async def dispatch(self, request: Request, call_next: Callable) -> JSONResponse:
        """
        中间件处理逻辑

        认证规则：
        1. 如果提供了 Authorization 头，必须通过 SSO 认证（如果SSO已启用）
        2. 如果提供了 x-api-key 头，必须通过 API Key 认证（如果API Key已启用）
        3. 如果同时提供了两种认证信息，只需通过其中一种即可
        4. 如果都未提供，返回 401
        5. 如果提供了某认证头但该认证方式未启用，返回 401
        """
        # 跳过认证路径
        if self._should_skip_auth(request.url.path):
            return await call_next(request)

        # 如果没有配置任何认证验证器，跳过认证
        if not self.sso_validator and not self.api_key_validator:
            return await call_next(request)

        # 获取两种认证头
        auth_header, api_key = await self._get_auth_headers(request)

        # 至少需要提供一种认证信息
        if not auth_header and not api_key:
            return JSONResponse(
                status_code=401,
                content={
                    "detail": "缺少认证信息",
                    "message": "请在请求头中提供 Authorization 或 x-api-key",
                },
            )

        # 认证结果跟踪
        auth_errors = []

        # SSO 认证（如果提供了 Authorization 头）
        if self.sso_validator:
            if not auth_header:
                auth_errors.append("Authorization 认证失败：缺少Authorization heard 信息")
            else:
                token = auth_header[7:] if auth_header.startswith("Bearer ") else auth_header
                sso_result = await self.sso_validator.verify(f"Bearer {token}")
                if not sso_result.is_authorized:
                    auth_errors.append(f"Authorization认证失败: {sso_result.error_message or '未通过'}")

        # API Key 认证（如果提供了 x-api-key 头）
        if self.api_key_validator:
            if not api_key:
                auth_errors.append("API Key认证失败: 缺少API Key ")
            else:
                api_result = await self.api_key_validator.verify(api_key)
                if not api_result.is_authorized:
                    auth_errors.append(f"API Key认证失败: {api_result.error_message or '无效'}")

        # 如果所有提供的认证方式都失败，返回错误信息
        if len(auth_errors) > 0:
            return JSONResponse(
                status_code=401,
                content={
                    "detail": "认证失败",
                    "message": " / ".join(auth_errors) if auth_errors else "未知认证错误",
                },
            )

        # 设置用户信息
        user_info = {}
        if auth_header and self.sso_validator:
            user_info["sso_authenticated"] = True
        if api_key and self.api_key_validator:
            user_info["api_key_authenticated"] = True

        request.state.auth_user = user_info if user_info else {"authenticated": True}
        return await call_next(request)


def create_auth_middleware():
    """
    创建认证中间件配置

    Returns:
        包含验证器的配置对象，如果未配置任何认证方式则返回 None
    """
    sso_validator = None
    api_key_validator = None

    # 初始化SSO验证器（如果启用了且配置了有效的验证URL）
    if settings.auth_sso_enabled and settings.auth_sso_verify_url and settings.auth_sso_verify_url.strip():
        sso_validator = SSOAuthValidator(
            verify_url=settings.auth_sso_verify_url,
            timeout=settings.auth_timeout
        )
        logger.info(f"SSO认证已启用，验证URL: {settings.auth_sso_verify_url}")
    elif settings.auth_sso_enabled:
        logger.warning("SSO认证已启用但未配置 verify_url，请设置 auth.sso.verify_url")
        raise ValueError("SSO认证已启用但未配置 verify_url，请设置 auth.sso.verify_url")

    # 初始化API Key验证器（如果启用了且配置了有效密钥）
    if settings.auth_api_key_enabled and settings.auth_api_key_valid_keys:
        api_key_validator = APIKeyValidator(
            valid_keys=settings.auth_api_key_valid_keys
        )
        logger.info(f"API Key认证已启用，有效密钥数: {len(settings.auth_api_key_valid_keys)}")

    # 如果没有配置任何认证方式，记录警告并返回 None
    if not sso_validator and not api_key_validator:
        logger.warning("未配置任何有效的认证方式（SSO或API Key），将无法处理请求")
        logger.warning("请在配置文件中设置 auth.sso.verify_url 或 auth.api_key.valid_keys")
        return None

    # 返回配置对象，而不是中间件实例
    return AuthMiddlewareConfig(
        sso_validator=sso_validator,
        api_key_validator=api_key_validator,
        skip_paths=settings.auth_skip_auth_paths,
    )
