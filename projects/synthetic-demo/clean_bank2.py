from pathlib import Path
from datetime import datetime, date
import re

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE_FOLDER = ROOT / "BANK2" / "source"
OUTPUT_FOLDER = ROOT / "BANK2" / "output"
OUTPUT_PREFIX = "BANK2_cleaned_"
DATE_CELL = "F2"
TRANSFER_PHRASE = "Перевод средств между своими счетами"


def safe_filename_part(value):
    if value is None or not str(value).strip():
        raise ValueError("Дата выписки отсутствует.")
    if isinstance(value, (datetime, date)):
        text = value.strftime("%d.%m.%Y")
    else:
        text = str(value).strip()
        for fmt in (
            "%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d",
            "%d/%m/%Y", "%Y%m%d", "%d.%m.%y",
        ):
            try:
                text = datetime.strptime(text, fmt).strftime("%d.%m.%Y")
                break
            except ValueError:
                continue
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text).strip(" .")
    if not text:
        raise ValueError("Дата выписки не подходит для имени файла.")
    return text


def find_cell(ws, phrase):
    target = phrase.strip().casefold()
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and target in str(cell.value).strip().casefold():
                return cell
    return None


def process_file(path):
    wb = load_workbook(path)
    try:
        if not wb.worksheets:
            raise ValueError("В книге нет листов.")
        date_part = safe_filename_part(wb.worksheets[0][DATE_CELL].value)

        for ws in wb.worksheets:
            rows_to_delete = set()

            incoming = find_cell(ws, "Входящее сальдо")
            if incoming and incoming.column < ws.max_column:
                ws.cell(incoming.row, incoming.column + 1).value = None

            for row in ws.iter_rows():
                if any(
                    cell.value is not None
                    and TRANSFER_PHRASE.casefold() in str(cell.value).casefold()
                    for cell in row
                ):
                    rows_to_delete.add(row[0].row)

            debit_header = find_cell(ws, "Дебет")
            if debit_header:
                for row_num in range(debit_header.row + 1, ws.max_row + 1):
                    value = ws.cell(row_num, debit_header.column).value
                    if value is not None and str(value).strip():
                        rows_to_delete.add(row_num)

            outgoing = find_cell(ws, "Исходящее сальдо")
            if outgoing:
                rows_to_delete.update(range(outgoing.row, ws.max_row + 1))

            for row_num in sorted(rows_to_delete, reverse=True):
                ws.delete_rows(row_num, 1)

        OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
        target = OUTPUT_FOLDER / f"{OUTPUT_PREFIX}{date_part}.xlsx"
        wb.save(target)
        print(f"BANK2: готово — {target.name}")
    finally:
        wb.close()


def main():
    if not SOURCE_FOLDER.exists():
        print(f"BANK2: папка не найдена: {SOURCE_FOLDER}")
        return
    files = sorted(
        path for path in SOURCE_FOLDER.glob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if not files:
        print("BANK2: исходные XLSX не найдены.")
        return
    for path in files:
        try:
            process_file(path)
        except Exception as exc:
            print(f"BANK2: ошибка {path.name}: {exc}")


if __name__ == "__main__":
    main()
