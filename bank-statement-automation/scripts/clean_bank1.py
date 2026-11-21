from pathlib import Path
import re
from datetime import date

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

ROOT = Path(__file__).resolve().parent.parent
SOURCE_FOLDER = ROOT / "test-data" / "BANK1" / "source"
OUTPUT_FOLDER = ROOT / "test-data" / "BANK1" / "output"
OUTPUT_PREFIX = "BANK1_cleaned_"


def find_cell(ws, phrase):
    phrase = phrase.casefold()
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            if phrase in str(cell.value).casefold():
                return cell
    return None


def merged_row_span(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.max_row
    return row, row


def expand_deletions_to_merged_rows(merged_ranges, rows_to_delete):
    rows = set(rows_to_delete)
    changed = True
    while changed:
        changed = False
        for rng in merged_ranges:
            if any(rng.min_row <= row <= rng.max_row for row in rows):
                for row in range(rng.min_row, rng.max_row + 1):
                    if row not in rows:
                        rows.add(row)
                        changed = True
    return rows


def physically_delete_rows_preserving_merges(ws, rows_to_delete):
    if not rows_to_delete:
        return
    original_merges = list(ws.merged_cells.ranges)
    rows_to_delete = expand_deletions_to_merged_rows(original_merges, rows_to_delete)

    for rng in original_merges:
        ws.unmerge_cells(str(rng))

    for row in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row, 1)

    for rng in original_merges:
        if any(row in rows_to_delete for row in range(rng.min_row, rng.max_row + 1)):
            continue
        removed_before_start = sum(row < rng.min_row for row in rows_to_delete)
        removed_before_end = sum(row < rng.max_row for row in rows_to_delete)
        ws.merge_cells(
            start_row=rng.min_row - removed_before_start,
            start_column=rng.min_col,
            end_row=rng.max_row - removed_before_end,
            end_column=rng.max_col,
        )


def extract_date_from_filename(path):
    patterns = [
        (r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{4})(?!\d)", "dmy"),
        (r"(?<!\d)(\d{4})[.\-_](\d{1,2})[.\-_](\d{1,2})(?!\d)", "ymd"),
        (r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)", "ymd"),
    ]
    for pattern, order in patterns:
        match = re.search(pattern, path.stem)
        if not match:
            continue
        a, b, c = map(int, match.groups())
        day, month, year = (a, b, c) if order == "dmy" else (c, b, a)
        try:
            return date(year, month, day).strftime("%d.%m.%Y")
        except ValueError:
            continue
    raise ValueError(f"Дата в имени файла не распознана: {path.name}")


def process_file(path):
    stamp = extract_date_from_filename(path)
    wb = load_workbook(path)
    try:
        for ws in wb.worksheets:
            debit_header = find_cell(ws, "Дебет / Debit")
            opening_label = find_cell(ws, "Входящий остаток на начало дня")
            operations_label = find_cell(ws, "Количество операций")

            if opening_label:
                row_num = opening_label.row + 1
                col_num = opening_label.column
                for rng in ws.merged_cells.ranges:
                    if (rng.min_row <= row_num <= rng.max_row
                            and rng.min_col <= col_num <= rng.max_col):
                        row_num, col_num = rng.min_row, rng.min_col
                        break
                cell = ws.cell(row_num, col_num)
                if not isinstance(cell, MergedCell):
                    cell.value = None

            rows_to_delete = set()
            phrase = "перевод собственных средств"
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell) or cell.value is None:
                        continue
                    if phrase in str(cell.value).casefold():
                        rows_to_delete.add(cell.row)
                        break

            if debit_header:
                _, header_bottom = merged_row_span(
                    ws, debit_header.row, debit_header.column
                )
                first_data_row = header_bottom + 2
                last_data_row = operations_label.row - 1 if operations_label else ws.max_row
                for row_num in range(first_data_row, last_data_row + 1):
                    cell = ws.cell(row_num, debit_header.column)
                    if (not isinstance(cell, MergedCell)
                            and cell.value is not None
                            and str(cell.value).strip()):
                        top, bottom = merged_row_span(
                            ws, row_num, debit_header.column
                        )
                        rows_to_delete.update(range(top, bottom + 1))

            if operations_label:
                rows_to_delete.update(
                    range(operations_label.row, ws.max_row + 1)
                )

            physically_delete_rows_preserving_merges(ws, rows_to_delete)

        OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
        target = OUTPUT_FOLDER / f"{OUTPUT_PREFIX}{stamp}.xlsx"
        wb.save(target)
        print(f"BANK1: готово — {target.name}")
    finally:
        wb.close()


def main():
    if not SOURCE_FOLDER.exists():
        print(f"BANK1: папка не найдена: {SOURCE_FOLDER}")
        return
    files = sorted(
        path for path in SOURCE_FOLDER.glob("*.xlsx")
        if not path.name.startswith("~$")
        and not path.name.startswith(OUTPUT_PREFIX)
    )
    if not files:
        print("BANK1: исходные XLSX не найдены.")
        return
    for path in files:
        try:
            process_file(path)
        except Exception as exc:
            print(f"BANK1: ошибка {path.name}: {exc}")


if __name__ == "__main__":
    main()
