import logging
from logging import Logger

# 记录日志模块是否已经初始化
_logging_initialized = False


def _setup_logging():
    """配置根日志器"""
    global _logging_initialized
    if _logging_initialized:
        return
    
    # 延迟导入 settings 以避免循环引用
    from src.config import settings
    
    # 设置根日志级别（与 config.toml 中的 debug 配置同步）
    root_logger = logging.getLogger()
    if settings.debug:
        root_logger.setLevel(logging.DEBUG)
    else:
        root_logger.setLevel(logging.INFO)
    
    # 确保有 Handler
    if not root_logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            "%(asctime)s %(levelname)-8s [%(name)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        handler.setFormatter(formatter)
        root_logger.addHandler(handler)
    else:
        # 更新已有 handler 的日志级别
        for handler in root_logger.handlers:
            handler.setLevel(logging.DEBUG if settings.debug else logging.INFO)
    
    _logging_initialized = True


def get_logger(name: str) -> Logger:
    """
    获取日志记录器。
    
    该函数会确保全局日志配置已初始化，并根据配置设置日志级别。
    """
    # 确保日志系统已初始化
    _setup_logging()
    
    logger = logging.getLogger(name)
    return logger


