# logger_setup.py
import logging
from logging.handlers import RotatingFileHandler
import os

# Шлях до файлу логу
path_to_log = "../logs/scrapper.log"
max_log_size = 10 * 1024 * 1024  # 10 MB

# Створення директорії для логів
log_dir = os.path.dirname(path_to_log)
os.makedirs(log_dir, exist_ok=True)

# Створюємо логер
logger = logging.getLogger("scrapper_logger")  # Ім'я логера
logger.setLevel(logging.INFO)

# Перевірка, чи вже додано хендлери (щоб уникнути дублікатів)
if not logger.handlers:
    # Створюємо обробник для файлу
    handler = RotatingFileHandler(path_to_log, maxBytes=max_log_size, backupCount=1)
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    # Обробник для консолі (опціонально)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
