"""Очистка обезличенного формата выписки БАНКА 2 (.xlsx).
Исходник не перезаписывается.
"""
from datetime import date, datetime
from pathlib import Path
import re
from openpyxl import load_workbook


def find_cell(ws, phrase):
    target = phrase.strip().casefold()
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and target in str(cell.value).strip().casefold():
                return cell
    return None


def normalize_date(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y")
    raw = str(value).strip() if value is not None else ""
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y%m%d", "%d.%m.%y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d.%m.%Y")
        except ValueError:
            pass
    raw = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", raw).strip(" .")
    if not raw:
        raise ValueError("Дата в ячейке F2 отсутствует.")
    return raw


def clean_file(source, output_dir):
    wb = load_workbook(source)
    try:
        if not wb.worksheets:
            raise ValueError("В книге нет листов.")
        stamp = normalize_date(wb.worksheets[0]["F2"].value)

        for ws in wb.worksheets:
            rows = set()
            incoming = find_cell(ws, "Входящее сальдо")
            if incoming and incoming.column < ws.max_column:
                ws.cell(incoming.row, incoming.column + 1).value = None

            for row in ws.iter_rows():
                if any(cell.value is not None and
                       "перевод средств между своими счетами" in str(cell.value).casefold()
                       for cell in row):
                    rows.add(row[0].row)

            debit = find_cell(ws, "Дебет")
            if debit:
                for r in range(debit.row + 1, ws.max_row + 1):
                    value = ws.cell(r, debit.column).value
                    if value is not None and str(value).strip():
                        rows.add(r)

            outgoing = find_cell(ws, "Исходящее сальдо")
            if outgoing:
                rows.update(range(outgoing.row, ws.max_row + 1))

            for r in sorted(rows, reverse=True):
                ws.delete_rows(r, 1)

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        target = output_dir / f"bank_2_{stamp}.xlsx"
        wb.save(target)
        return target
    finally:
        wb.close()


def process_folder(input_dir, output_dir):
    folder = Path(input_dir)
    if not folder.exists():
        print(f"БАНК 2: папка не найдена: {folder}")
        return 0
    count = 0
    for source in sorted(folder.glob("*.xlsx")):
        if source.name.startswith("~$"):
            continue
        try:
            target = clean_file(source, output_dir)
            print(f"БАНК 2: готово — {target.name}")
            count += 1
        except Exception as exc:
            print(f"БАНК 2: ошибка {source.name}: {exc}")
    return count
