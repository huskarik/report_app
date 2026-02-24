import logging
import sys


def setup_logger(name: str = "app") -> logging.Logger:
    logger = logging.getLogger(name)

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    # вывод в консоль
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)

    logger.addHandler(console_handler)

    return logger
