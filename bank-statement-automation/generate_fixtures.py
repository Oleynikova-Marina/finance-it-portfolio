from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent
DATE = datetime(2026, 11, 14)


def create_bank1() -> None:
    folder = ROOT / "BANK1" / "source"
    folder.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Выписка"
    ws.merge_cells("A1:E1")
    ws["A1"] = "БАНК 1"
    ws.merge_cells("A2:E2")
    ws["A2"] = "СИНТЕТИЧЕСКАЯ ВЫПИСКА • ДЕМОНСТРАЦИОННЫЕ ДАННЫЕ"
    ws["A4"] = "Дата формирования"
    ws["B4"] = DATE
    ws["A6"] = "Входящий остаток на начало дня"
    ws["A7"] = 1_250_000

    headers = ["Дата", "Номер документа", "Назначение платежа", "Дебет / Debit", "Кредит / Credit"]
    for col, value in enumerate(headers, start=1):
        ws.cell(14, col, value)
        ws.cell(15, col, "Показатели")
        ws.merge_cells(start_row=14, start_column=col, end_row=15, end_column=col)
        cell = ws.cell(14, col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Row 16 is intentionally left blank to mimic the source layout.
    transactions = [
        [DATE, "1001", "Оплата по договору поставки № 15", None, 185_000],
        [DATE, "1002", "Комиссия банка за обслуживание", 1_250, None],
        [DATE, "1003", "Перевод собственных средств между счетами", None, 50_000],
        [DATE, "1004", "Поступление оплаты от покупателя по договору № 22", None, 320_000],
        [DATE, "1005", "Оплата услуг связи", 890, None],
    ]
    for row_num, values in enumerate(transactions, start=17):
        for col_num, value in enumerate(values, start=1):
            if value is not None:
                ws.cell(row_num, col_num, value)

    ws["A22"], ws["B22"] = "Количество операций", 5
    ws["A23"], ws["B23"] = "Итого дебет", 2_140
    ws["A24"], ws["B24"] = "Итого кредит", 555_000
    for col, width in {"A": 16, "B": 18, "C": 55, "D": 18, "E": 18}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A17"
    ws.sheet_view.showGridLines = False
    wb.save(folder / "BANK1_statement_14.11.2026.xlsx")
    wb.close()


def create_bank2() -> None:
    folder = ROOT / "BANK2" / "source"
    folder.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Выписка"
    ws["A1"] = "БАНК 2"
    ws["A2"] = "СИНТЕТИЧЕСКАЯ ВЫПИСКА — ВСЕ ДАННЫЕ ВЫМЫШЛЕНЫ"
    ws["E2"], ws["F2"] = "Дата выписки", DATE
    ws["A4"], ws["B4"] = "Входящее сальдо", 875_000
    headers = ["Дата", "Документ", "Плательщик / получатель", "Дебет", "Кредит", "Код контрагента", "Банк", "Назначение платежа"]
    for col, value in enumerate(headers, start=1):
        cell = ws.cell(7, col, value)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    transactions = [
        [DATE, "ПЛ-2001", "ООО Альфа-Тест", None, 240_000, "TEST-001", "БАНК 2", "Оплата по договору № 31"],
        [DATE, "ПЛ-2002", "ООО Бета-Тест", 45_000, None, "TEST-002", "БАНК 2", "Оплата услуг по договору № 08"],
        [DATE, "ПЛ-2003", "ООО Гамма-Тест", None, 125_000, "TEST-003", "БАНК 2", "Перевод средств между своими счетами"],
        [DATE, "ПЛ-2004", "ООО Дельта-Тест", None, 76_000, "TEST-004", "БАНК 2", "Поступление оплаты за услуги"],
        [DATE, "ПЛ-2005", "ООО Эпсилон-Тест", 3_200, None, "TEST-005", "БАНК 2", "Комиссия за обслуживание счёта"],
    ]
    for row_num, values in enumerate(transactions, start=8):
        for col_num, value in enumerate(values, start=1):
            if value is not None:
                ws.cell(row_num, col_num, value)

    ws["A13"], ws["B13"] = "Исходящее сальдо", 1_267_800
    ws["A14"] = "Строка после исходящего сальдо — должна быть удалена"
    ws["A15"] = "Дополнительная служебная информация"
    widths = [16, 16, 25, 14, 14, 18, 16, 42]
    for col_num, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + col_num)].width = width
    ws.freeze_panes = "A8"
    ws.sheet_view.showGridLines = False
    wb.save(folder / "BANK2_statement_14.11.2026.xlsx")
    wb.close()


def main() -> None:
    create_bank1()
    create_bank2()
    print("Созданы синтетические исходные выписки BANK1 и BANK2.")


if __name__ == "__main__":
    main()
