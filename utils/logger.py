# utils/logger.py

import logging

class logger:
    """Simple logging wrapper."""

    @staticmethod
    def setup():
        logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    @staticmethod
    def info(msg: str):
        logging.info(msg)

    @staticmethod
    def warning(msg: str):
        logging.warning(msg)

    @staticmethod
    def error(msg: str):
        logging.error(msg)
