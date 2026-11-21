# -*- coding: utf-8 -*-
"""Создаёт обезличенные Excel-файлы для тестирования проекта 02.

Запуск из корня репозитория:
    py debt-payment-reconciliation/test-data/generate_test_data.py

Файлы создаются в debt-payment-reconciliation/scripts/{input,working,data}.
Существующие файлы с такими именами будут перезаписаны.
"""
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

PROJECT_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = PROJECT_DIR / "scripts"

HEADERS = [
    "№", "Юрлицо", "Код", "Дата оплаты по плану до", "Контрагент",
    "Документ", "Статья", "Сумма плана", "Валюта", "Количество",
    "Примечание", "Долг по отгр", "Группа", "Город",
]

ROWS = [
    [1, "КОМПАНИЯ 1", "C-001", datetime(2026, 11, 20), "Поставщик 1",
     "Счёт Т-001", "Поставщики", 150000, "RUB", 1, "Часть группы", 100000, "G-MSK", "МСК"],
    [2, "КОМПАНИЯ 1", "C-001", datetime(2026, 11, 20), "Поставщик 1",
     "Счёт Т-002", "Поставщики", 150000, "RUB", 1, "Вторая строка той же группы", 50000, "G-MSK", "МСК"],
    [3, "КОМПАНИЯ 1", "C-002", datetime(2026, 11, 21), "Поставщик Бета-Тест",
     "Счёт Т-003", "Поставщики", 72000, "RUB", 1, "Два подходящих платежа — неоднозначность", 72000, "G-SPB", "САНКТ-ПЕТЕРБУРГ"],
    [4, "КОМПАНИЯ 1", "C-003", datetime(2026, 11, 22), "Поставщик Гамма-Тест",
     "Счёт Т-004", "Поставщики", 39000, "RUB", 1, "Та же сумма, другой город", 39000, "G-EKB", "ЕКБ"],
    [5, "КОМПАНИЯ 1", "C-004", datetime(2026, 11, 23), "Поставщик Дельта-Тест",
     "Счёт Т-005", "Поставщики", 24500, "RUB", 1, "Подходящего платежа нет", 24500, "G-KZN", "КАЗАНЬ"],
    [6, "КОМПАНИЯ 2", "C-005", datetime(2026, 11, 20), "Поставщик Эпсилон-Тест",
     "Счёт Т-006", "Поставщики", 85000, "RUB", 1, "Сокращение города в назначении", 85000, "G-MSK2", "МОСКВА"],
    [7, "КОМПАНИЯ 2", "C-006", datetime(2026, 11, 24), "Поставщик Дзета-Тест",
     "Счёт Т-007", "Поставщики", 26000, "RUB", 1, "Проверка русской шапки выписки", 26000, "G-KZN2", "КАЗАНЬ"],
]


def style_header(ws, row=2):
    for cell in ws[row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="365F91")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = f"A{row + 1}"
    ws.sheet_view.showGridLines = False


def write_plan_sheet(ws):
    ws.cell(1, 1, "СИНТЕТИЧЕСКИЕ ТЕСТОВЫЕ ДАННЫЕ")
    for col, value in enumerate(HEADERS, 1):
        ws.cell(2, col, value)
    for row_idx, row in enumerate(ROWS, 3):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row_idx, col_idx, value)
    style_header(ws)
    for col in range(1, len(HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions["E"].width = 26
    ws.column_dimensions["K"].width = 42
    ws.column_dimensions["N"].width = 26


def save_source_and_working():
    input_dir = SCRIPTS_DIR / "input"
    working_dir = SCRIPTS_DIR / "working"
    input_dir.mkdir(parents=True, exist_ok=True)
    working_dir.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Источник"
    write_plan_sheet(ws)
    wb.save(input_dir / "Исходные_данные_тест.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws.cell(1, 1, "Рабочая книга для теста")
    for col, value in enumerate(HEADERS, 1):
        ws.cell(2, col, value)
    style_header(ws)
    wb.save(working_dir / "Рабочая_книга_тест.xlsx")


def save_plan():
    plan_dir = SCRIPTS_DIR / "data" / "plan"
    plan_dir.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    write_plan_sheet(ws)
    send = wb.create_sheet("send")
    for col, value in enumerate(HEADERS, 1):
        send.cell(2, col, value)
    style_header(send)
    wb.create_sheet("total")
    wb.save(plan_dir / "Report_20.11.2026_ТЕСТ.xlsx")


def save_bank(folder, filename, headers, operations):
    out_dir = SCRIPTS_DIR / "data" / folder
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Операции"
    for col, value in enumerate(headers, 1):
        ws.cell(1, col, value)
    for row_idx, row in enumerate(operations, 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row_idx, col_idx, value)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    wb.save(out_dir / filename)


def main():
    save_source_and_working()
    save_plan()

    bilingual_headers = [
        "Дата операции/Posting date",
        "Дебет / Debit",
        "Реквизиты корреспондента / Counter party details",
        "Основание операции (назначение платежа) / Payment details",
    ]
    bank_1 = [
        [datetime(2026, 11, 20), 150000, "ИНН 9900000001", "Оплата по счетам Т-001, Т-002. Москва"],
        [datetime(2026, 11, 20), 72000, "ИНН 9900000002", "Оплата по счету Т-003. СПБ"],
        [datetime(2026, 11, 20), 72000, "ИНН 9900000002", "Повторная операция по счету Т-003. СПБ"],
        [datetime(2026, 11, 20), 39000, "ИНН 9900000003", "Оплата по счету Т-004. Казань"],
    ]
    save_bank("bank_1", "Выписка_Банк_1_ТЕСТ.xlsx", bilingual_headers, bank_1)

    russian_headers = ["Дата операции", "Дебет", "ИНН контрагента", "Назначение платежа"]
    bank_2 = [
        [datetime(2026, 11, 20), 85000, "ИНН 9900000001", "Оплата по счету Т-006. МСК"],
        [datetime(2026, 11, 20), 26000, "ИНН 9900000002", "Оплата по счету Т-007. Казань"],
    ]
    save_bank("bank_2", "Выписка_Банк_2_ТЕСТ.xlsx", russian_headers, bank_2)

    print("Тестовые книги созданы в:", SCRIPTS_DIR)
    print("Контрольная сумма исходного столбца L: 396 500,00")
    print("Для reconcile_bank_payments.py задайте COUNTERPARTY_INNS = "
          '("9900000001", "9900000002", "9900000003")')


if __name__ == "__main__":
    main()
