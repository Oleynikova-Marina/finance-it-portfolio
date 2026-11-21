"""Локальный пример настроек. Скопируйте в config.py и заполните у себя.
Не публикуйте config.py: он может содержать личные пути и адреса.
"""
from pathlib import Path

BASE_DIR = Path.home() / "Documents" / "finance_automation"
BANK_1_INPUT = BASE_DIR / "input" / "bank_1"
BANK_2_INPUT = BASE_DIR / "input" / "bank_2"
OUTPUT_DIR = BASE_DIR / "output"

OUTLOOK_SENDER = "sender@example.com"
OUTLOOK_RECIPIENTS = ["recipient@example.com"]
SIGNATURE_NAME = "Ваше имя"
SIGNATURE_PHONE = "Ваш телефон"
