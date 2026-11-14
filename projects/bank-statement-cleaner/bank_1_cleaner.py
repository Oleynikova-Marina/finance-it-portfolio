"""Очистка обезличенного формата выписки БАНКА 1 (.xlsx).
Исходник не перезаписывается. Поддерживаются объединённые ячейки.
"""
from datetime import date
from pathlib import Path
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


def find_cell(ws, phrase):
    phrase = phrase.casefold()
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell, MergedCell) and cell.value is not None and phrase in str(cell.value).casefold():
                return cell
    return None


def merged_span(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.max_row
    return row, row


def delete_rows_keep_merges(ws, rows):
    if not rows:
        return
    ranges = list(ws.merged_cells.ranges)
    rows = set(rows)
    changed = True
    while changed:
        changed = False
        for rng in ranges:
            if any(rng.min_row <= r <= rng.max_row for r in rows):
                for r in range(rng.min_row, rng.max_row + 1):
                    if r not in rows:
                        rows.add(r)
                        changed = True
    for rng in ranges:
        ws.unmerge_cells(str(rng))
    for r in sorted(rows, reverse=True):
        ws.delete_rows(r, 1)
    for rng in ranges:
        if any(r in rows for r in range(rng.min_row, rng.max_row + 1)):
            continue
        shift_a = sum(r < rng.min_row for r in rows)
        shift_b = sum(r < rng.max_row for r in rows)
        ws.merge_cells(start_row=rng.min_row-shift_a, start_column=rng.min_col,
                       end_row=rng.max_row-shift_b, end_column=rng.max_col)


def filename_date(path):
    for pattern, order in [
        (r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{4})(?!\d)", "dmy"),
        (r"(?<!\d)(\d{4})[.\-_](\d{1,2})[.\-_](\d{1,2})(?!\d)", "ymd"),
        (r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)", "ymd"),
    ]:
        m = re.search(pattern, path.stem)
        if m:
            a, b, c = map(int, m.groups())
            d, mo, y = (a,b,c) if order == "dmy" else (c,b,a)
            try:
                return date(y,mo,d).strftime("%d.%m.%Y")
            except ValueError:
                pass
    raise ValueError("Дата в имени исходного файла не распознана.")


def clean_file(source, output_dir):
    stamp = filename_date(source)
    wb = load_workbook(source)
    try:
        for ws in wb.worksheets:
            debit = find_cell(ws, "Дебет / Debit")
            opening = find_cell(ws, "Входящий остаток на начало дня")
            summary = find_cell(ws, "Количество операций")
            if opening:
                r, c = opening.row + 1, opening.column
                for rng in ws.merged_cells.ranges:
                    if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                        r, c = rng.min_row, rng.min_col
                        break
                cell = ws.cell(r,c)
                if not isinstance(cell, MergedCell):
                    cell.value = None

            rows = set()
            for row in ws.iter_rows():
                if any(not isinstance(cell, MergedCell) and cell.value is not None
                       and "перевод собственных средств" in str(cell.value).casefold()
                       for cell in row):
                    rows.add(row[0].row)

            if debit:
                _, header_bottom = merged_span(ws, debit.row, debit.column)
                first = header_bottom + 2
                last = summary.row - 1 if summary else ws.max_row
                for r in range(first, last + 1):
                    cell = ws.cell(r, debit.column)
                    if not isinstance(cell, MergedCell) and cell.value is not None and str(cell.value).strip():
                        a,b = merged_span(ws,r,debit.column)
                        rows.update(range(a,b+1))
            if summary:
                rows.update(range(summary.row, ws.max_row+1))
            delete_rows_keep_merges(ws, rows)

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        target = output_dir / f"bank_1_{stamp}.xlsx"
        wb.save(target)
        return target
    finally:
        wb.close()


def process_folder(input_dir, output_dir):
    folder = Path(input_dir)
    if not folder.exists():
        print(f"БАНК 1: папка не найдена: {folder}")
        return 0
    count = 0
    for source in sorted(folder.glob("*.xlsx")):
        if source.name.startswith("~$"):
            continue
        try:
            target = clean_file(source, output_dir)
            print(f"БАНК 1: готово — {target.name}")
            count += 1
        except Exception as exc:
            print(f"БАНК 1: ошибка {source.name}: {exc}")
    return count
