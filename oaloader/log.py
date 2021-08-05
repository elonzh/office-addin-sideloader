import sys

from loguru import logger

from .const import app_dir


def setup_logger(level: str):
    level = level.upper()
    logger.remove()
    fmt = (
        "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green>"
        "[<level>{level: <8}</level>]"
        "<level>{message}</level> - "
        "{extra}"
    )
    logger.add(
        sys.stdout,
        level=level,
        format=fmt,
    )
    logger.add(
        app_dir.joinpath("main.log"),
        level="DEBUG",
        format=fmt,
    )
