# -*- coding: utf-8 -*-
"""Сверка плана оплат с банковскими выписками.

Группа переносится на лист send при совпадении суммы и города в назначении платежа;
дата операции банка может отличаться от плановой даты. Перенесённые строки
физически удаляются с листа data. Перед сохранением создаётся резервная копия.
"""
from __future__ import annotations

import re
import shutil
import warnings
import tempfile
from collections import defaultdict
from copy import copy
from datetime import datetime, date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import openpyxl
import win32com.client
from openpyxl.cell.cell import MergedCell

# Убираем только служебное предупреждение о стиле исходной банковской книги.
warnings.filterwarnings("ignore", message=r"Workbook contains no default style.*")

# -------------------------- НАСТРОЙКИ ---------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
BANK_DIRS = (
    DATA_DIR / "bank_1",
    DATA_DIR / "bank_2",
)
PLAN_DIR = DATA_DIR / "plan"
PLAN_SHEET = "data"
RESULT_SHEET = "send"
TOTAL_SHEET = "total"
TREEMAP_SHEET = "treemap"  # существующий лист; сохраняем отдельно от total

# Проверяем заданные юрлица по выпискам из папок bank_1 и bank_2.
# Замените примеры на условные названия организаций из вашей тестовой книги.
LEGAL_ENTITY_FILTERS = ("КОМПАНИЯ 1", "КОМПАНИЯ 2")

# ИНН контрагентов для демонстрационного набора из test-data/generate_test_data.py.
# Это вымышленные значения; при работе со своими тестовыми файлами замените их
# на соответствующие синтетические ИНН. Не добавляйте реальные реквизиты в публичный код.
COUNTERPARTY_INNS = ("9900000001", "9900000002", "9900000003")

# Для данной выписки заголовки расположены в строках 15–17,
# данные начинаются со строки 18. Скрипт ищет заголовки автоматически.
BANK_HEADER_SCAN_ROWS = 25
AMOUNT_TOLERANCE = Decimal("0.01")  # точность сравнения до копейки
# ---------------------------------------------------------------------------


def norm(value) -> str:
    """Текст для нечувствительного к регистру сравнения."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def entity_key(value) -> str:
    """Нормализует название юрлица: убирает кавычки и типовые приставки ООО/АО."""
    text = norm(value)
    text = re.sub(r'[\"«»„“”]', "", text)
    text = re.sub(r"\b(ооо|ао|пао|зао|оао)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def same_entity(left, right) -> bool:
    """Сопоставляет юрлица по нормализованному названию, даже если есть приставка ООО."""
    a, b = entity_key(left), entity_key(right)
    return bool(a and b and (a == b or a in b or b in a))


def digits(value) -> str:
    return re.sub(r"\D", "", str(value or ""))


def money(value) -> Decimal | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, Decimal):
        amount = value
    elif isinstance(value, (int, float)):
        amount = Decimal(str(value))
    else:
        s = str(value).replace("\u00a0", " ").replace("руб.", "").replace("₽", "").strip()
        s = s.replace(" ", "").replace(",", ".")
        try:
            amount = Decimal(s)
        except InvalidOperation:
            return None
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def date_key(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    s = str(value).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None


def find_header_row(ws, required_headers, scan_rows=30):
    required = {norm(x) for x in required_headers}
    for r in range(1, min(ws.max_row, scan_rows) + 1):
        found = set()
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None:
                found.add(norm(v))
        if required.issubset(found):
            return r
    return None


def header_columns(ws, header_row):
    result = {}
    for c in range(1, ws.max_column + 1):
        value = ws.cell(header_row, c).value
        if value is not None:
            result[norm(value)] = c
    return result


def find_col(columns, *names):
    for name in names:
        key = norm(name)
        if key in columns:
            return columns[key]
    return None


def cell_value(ws, row, col):
    if not col:
        return None
    return ws.cell(row, col).value


def city_from_rs(value):
    """Нормализует город из поля рс: удаляет служебные маркеры ОБЩИЙ и СОСЕНКИ."""
    city = str(value or "").replace("\u00a0", " ").strip()
    city = re.sub(r"^\s*ОБЩИЙ\s*[,;:—–-]?\s*", "", city, flags=re.IGNORECASE)
    city = re.sub(r"(?:\s*[-–—]\s*|\s+)СОСЕНКИ\s*$", "", city, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", city).strip(" ,;:-—–").upper()


def total_city_from_rs(value):
    """Возвращает нормализованное название города для итогового листа."""
    return city_from_rs(value)


def date_from_filename(path: Path):
    """Ищет дату в имени файла: ДД.ММ.ГГГГ, ДД.ММ.ГГ или ГГГГ-ММ-ДД."""
    stem = path.stem
    patterns = [
        (r"(?<!\d)(\d{2}\.\d{2}\.\d{4})(?!\d)", "%d.%m.%Y"),
        (r"(?<!\d)(\d{2}\.\d{2}\.\d{2})(?!\d)", "%d.%m.%y"),
        (r"(?<!\d)(\d{4}-\d{2}-\d{2})(?!\d)", "%Y-%m-%d"),
    ]
    for pattern, fmt in patterns:
        match = re.search(pattern, stem)
        if match:
            try:
                return datetime.strptime(match.group(1), fmt).date()
            except ValueError:
                continue
    raise RuntimeError(f"Не удалось определить дату из названия файла: {path.name}. Используй дату в формате ДД.ММ.ГГ или ДД.ММ.ГГГГ.")


def update_total_sheet(wb, data_ws, c_rs, c_debt, header_row, report_date):
    """Создаёт отдельный лист total; treemap не переименовывает и не изменяет."""
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    # Важно: никогда не переименовываем другие листы в treemap или total.
    # Лист total создаётся/обновляется только по точному имени TOTAL_SHEET.

    # Создаём total, если его ещё нет. Никогда не используем treemap как total.
    if TOTAL_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(TOTAL_SHEET)
        history_cols = []
        old_cities = {}
    else:
        ws = wb[TOTAL_SHEET]
        old_dates = {}
        for c in range(2, ws.max_column + 1):
            d = date_key(ws.cell(2, c).value)
            if d:
                old_dates[c] = d
        old_cities = {}
        for r in range(3, ws.max_row + 1):
            city = total_city_from_rs(ws.cell(r, 1).value)
            if city and city.upper() not in {'ГОРОД', 'ИТОГО', 'ДАННЫЕ ПО ЗАДОЛЖЕННОСТИ ПО ГОРОДАМ, ТЫС.РУБ.'}:
                old_cities[r] = city
        looks_like_total = (
            str(ws.cell(1, 1).value or '').strip().casefold() == 'данные по городам, тыс.руб.'
            and bool(old_dates) and bool(old_cities)
        )
        history_cols = []
        if looks_like_total:
            for c, d in old_dates.items():
                values = {}
                for r, city in old_cities.items():
                    value = ws.cell(r, c).value
                    if isinstance(value, (int, float, Decimal)):
                        values[city.casefold()] = value
                history_cols.append((d, values))
        else:
            # Не переносим цветную визуализацию/чужую структуру в total.
            history_cols = []
            old_cities = {}
        for merged_range in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged_range))
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)
        for col in list(ws.column_dimensions):
            del ws.column_dimensions[col]

    totals = defaultdict(lambda: Decimal('0.00'))
    for r in range(header_row + 1, data_ws.max_row + 1):
        city = total_city_from_rs(data_ws.cell(r, c_rs).value)
        amount = money(data_ws.cell(r, c_debt).value)
        if city and amount is not None:
            totals[city] += amount

    all_cities = {city for city in totals if city}
    all_cities.update(city for city in old_cities.values() if city)

    # Если эта дата уже есть, обновляем её столбец; новую дату добавляем справа.
    # Поэтому повторный запуск за один день не создаёт дубли, а завтра появится новый столбец.
    date_columns = list(history_cols)
    matching_index = next((i for i, (dt, _) in enumerate(date_columns) if dt == report_date), None)
    if matching_index is None:
        date_columns.append((report_date, {}))
        current_index = len(date_columns) - 1
    else:
        current_index = matching_index
        date_columns[current_index] = (report_date, {})

    priority = {'МОСКВА': 0, 'СПБ': 1}
    cities_sorted = sorted(all_cities, key=lambda x: (priority.get(x.upper(), 2), x.casefold()))

    ws.cell(1, 1).value = 'ДАННЫЕ ПО ЗАДОЛЖЕННОСТИ ПО ГОРОДАМ, тыс.руб.'
    ws.cell(2, 1).value = 'Город'
    for col, (dt, _) in enumerate(date_columns, start=2):
        ws.cell(2, col).value = dt
        ws.cell(2, col).number_format = 'dd.mm.yyyy'

    for row, city in enumerate(cities_sorted, start=3):
        ws.cell(row, 1).value = city
        for col, (_, values) in enumerate(date_columns, start=2):
            if col - 2 == current_index:
                # Формула Excel пересчитывается при ручном удалении/изменении строк на data.
                # Поиск города через wildcard поддерживает значения рс вида "ОБЩИЙ, ЕКАТЕРИНБУРГ".
                rs_letter = get_column_letter(c_rs)
                debt_letter = get_column_letter(c_debt)
                city_cell = f'$A{row}'
                # SUMPRODUCT avoids the #VALUE! issue that can occur with whole-column
                # SUMIF ranges in some Excel workbooks. IFERROR also ignores non-numeric
                # cells in the amount column. Bounded ranges keep recalculation efficient.
                rs_range = f'data!${rs_letter}$2:${rs_letter}$10000'
                debt_range = f'data!${debt_letter}$2:${debt_letter}$10000'
                ws.cell(row, col).value = (
                    f'=ROUND(SUMIF({rs_range},"*"&{city_cell}&"*",{debt_range})/1000,0)'
                )
            elif city.casefold() in values:
                ws.cell(row, col).value = values[city.casefold()]

    # Строка ИТОГО всегда идёт сразу под последним городом; суммируем каждый столбец даты.
    total_row = len(cities_sorted) + 3
    ws.cell(total_row, 1).value = 'ИТОГО'
    for col in range(2, len(date_columns) + 2):
        # И текущий, и исторические столбцы суммируются формулой; текущий столбец
        # реагирует на ручное удаление строк на data без повторного запуска Python.
        from openpyxl.utils import get_column_letter
        letter = get_column_letter(col)
        ws.cell(total_row, col).value = f'=SUM({letter}3:{letter}{total_row - 1})'

    from openpyxl.styles import Font, Alignment, Border, Side
    for row in ws.iter_rows(min_row=1, max_row=max(2, total_row), min_col=1, max_col=max(2, len(date_columns) + 1)):
        for cell in row:
            cell.font = Font(name='Times New Roman', size=12, bold=True)
            cell.alignment = Alignment(vertical='center')
            if cell.row >= 3 and cell.column >= 2:
                cell.number_format = '# ##0;[Red]-# ##0'
    # Выделяем строку ИТОГО верхней границей и сохраняем числовой формат.
    total_border = Border(top=Side(style='thin'))
    for col in range(1, len(date_columns) + 2):
        total_cell = ws.cell(total_row, col)
        total_cell.font = Font(name='Times New Roman', size=12, bold=True)
        total_cell.border = total_border
        if col >= 2:
            total_cell.number_format = '# ##0;[Red]-# ##0'

    ws.column_dimensions['A'].width = max(18, min(40, max([len(str(c)) for c in cities_sorted] + [5, len('ИТОГО')]) + 3))
    for c in range(2, len(date_columns) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 16
    ws.freeze_panes = 'B3'
    ws.sheet_view.showGridLines = False
    return len(totals), len(date_columns) + 1


def city_matches(rs_value, payment_purpose):
    city = norm(city_from_rs(rs_value))
    purpose = norm(payment_purpose)
    if not city or not purpose:
        return False
    # Основное правило — название города из рс должно содержаться в назначении.
    if city in purpose:
        return True
    # Для составных названий допускаем совпадение по значимым частям:
    # например, «МОСКВА-СОСЕНКИ» может быть обозначено как «Оплата Москва».
    parts = [p for p in re.split(r"[\s\-–—,/]+", city) if len(p) >= 3]
    return bool(parts and any(p in purpose for p in parts))


def copy_row_values_and_style(src_ws, src_row, dst_ws, dst_row, max_col):
    for col in range(1, max_col + 1):
        source = src_ws.cell(src_row, col)
        target = dst_ws.cell(dst_row, col)
        target.value = source.value
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)
        if source.comment:
            target.comment = copy(source.comment)
    if src_ws.row_dimensions[src_row].height:
        dst_ws.row_dimensions[dst_row].height = src_ws.row_dimensions[src_row].height


def copy_sheet_layout(source_ws, target_ws, max_col):
    for col_letter, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = dim.width
        target_ws.column_dimensions[col_letter].hidden = dim.hidden
    target_ws.freeze_panes = source_ws.freeze_panes
    target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines
    for row in range(1, source_ws.max_row + 1):
        if source_ws.row_dimensions[row].height:
            target_ws.row_dimensions[row].height = source_ws.row_dimensions[row].height


def find_plan_file():
    candidates = [p for p in PLAN_DIR.iterdir()
                  if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
                  and not p.name.startswith("~$")]
    if len(candidates) != 1:
        raise RuntimeError(
            f"В папке {PLAN_DIR} должен быть ровно один файл Excel; найдено: {len(candidates)}."
        )
    return candidates[0]


def load_bank_rows(bank_file):
    """Читает два варианта структуры выписок: многострочную и обычную шапку."""
    keep_vba = bank_file.suffix.lower() == ".xlsm"
    wb = openpyxl.load_workbook(bank_file, data_only=True, read_only=False, keep_vba=keep_vba)
    found = []
    try:
        for ws in wb.worksheets:
            # Надёжнее определять юрлицо по папке выписки: в одном из форматов ячейка A3
            # может содержать заголовок отчёта, а не название организации.
            folder_name = bank_file.parent.name.casefold()
            if folder_name == "bank_1":
                bank_entity = "КОМПАНИЯ 1"
            elif folder_name == "bank_2":
                bank_entity = "КОМПАНИЯ 2"
            else:
                bank_entity = ws["A3"].value
                if bank_entity is None:
                    bank_entity = wb.worksheets[0]["A3"].value
            bank_entity = str(bank_entity).strip() if bank_entity is not None else ""

            header_row = find_header_row(
                ws,
                ["Дата операции/Posting date", "Дебет / Debit", "Реквизиты корреспондента / Counter party details",
                 "Основание операции (назначение платежа) / Payment details"],
                BANK_HEADER_SCAN_ROWS,
            )
            # Универсальное распознавание обычной и многострочной шапки.
            if header_row is None:
                for r in range(1, min(ws.max_row, BANK_HEADER_SCAN_ROWS) + 1):
                    row_text = " | ".join(norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1))
                    has_debit = "дебет" in row_text
                    has_purpose = any(term in row_text for term in ("назначение", "основание операции", "payment details"))
                    has_date = "дата" in row_text or "posting date" in row_text
                    has_inn = "инн" in row_text or "counter party details" in row_text or "реквизиты корреспондента" in row_text
                    if has_debit and has_purpose and (has_date or has_inn):
                        header_row = r
                        break
            if header_row is None:
                # Запасной вариант для выписки с нумерованными колонками.
                for r in range(1, min(ws.max_row, BANK_HEADER_SCAN_ROWS) + 1):
                    vals = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
                    if any(v in ("9", "дебет", "дебет / debit") for v in vals) and ws.max_column >= 23:
                        header_row = r
                        break
            if header_row is None:
                print(f"  [ПРОПУСК] Не удалось определить заголовки: {bank_file.name}, лист {ws.title}")
                continue

            col_date = col_debit = col_inn = col_purpose = None
            for r in range(max(1, header_row - 3), min(ws.max_row, header_row + 1) + 1):
                for c in range(1, ws.max_column + 1):
                    value = norm(ws.cell(r, c).value)
                    if value in ("дата", "дата операции", "posting date") or "дата операции" in value or "posting date" in value:
                        col_date = c
                    if value in ("дебет / debit", "дебет") or ("дебет" in value and col_debit is None):
                        col_debit = c
                    if ("инн контрагента" in value or "реквизиты корреспондента" in value
                            or "counter party details" in value):
                        col_inn = c
                    if (value == "назначение" or "назначение платежа" in value
                            or "основание операции" in value or "payment details" in value):
                        col_purpose = c

            # Фиксированные позиции — только резерв для прежнего формата выписки.
            if col_date is None and ws.max_column >= 3 and ws.max_column >= 23:
                col_date = 3
            if col_inn is None and ws.max_column >= 23:
                col_inn = 13
            if col_debit is None and ws.max_column >= 23:
                col_debit = 17
            if col_purpose is None and ws.max_column >= 23:
                col_purpose = 23

            if not all((col_date, col_debit, col_inn, col_purpose)):
                print(f"  [ПРОПУСК] Не найдены нужные столбцы: {bank_file.name}, лист {ws.title}")
                continue

            for r in range(header_row + 1, ws.max_row + 1):
                d = date_key(cell_value(ws, r, col_date))
                debit = money(cell_value(ws, r, col_debit))
                inn_text = str(cell_value(ws, r, col_inn) or "")
                purpose = cell_value(ws, r, col_purpose)
                if d is None or debit is None or not purpose:
                    continue
                inn_digits = digits(inn_text)
                if not any(inn in inn_digits for inn in COUNTERPARTY_INNS):
                    continue
                found.append({
                    "file": bank_file.name,
                    "sheet": ws.title,
                    "row": r,
                    "date": d,
                    "debit": debit,
                    "inn_text": inn_text,
                    "purpose": str(purpose),
                    "entity": bank_entity,
                    "folder": bank_file.parent.name,
                })
    finally:
        wb.close()
    return found


def main():
    print("=" * 68)
    print("СВЕРКА ПЛАНА С БАНКОВСКИМИ ВЫПИСКАМИ")
    print("Правило: сумма и город должны совпасть; дата операции может отличаться.")
    print("Папки выписок:")
    for folder in BANK_DIRS:
        print(f"  - {folder}")
    print(f"Папка плана:   {PLAN_DIR}")
    print("=" * 68)
    missing_dirs = [folder for folder in BANK_DIRS if not folder.exists()]
    if missing_dirs:
        raise RuntimeError("Не найдены папки банковских выписок:\n" + "\n".join(map(str, missing_dirs)))
    if not PLAN_DIR.exists():
        raise RuntimeError(f"Не найдена папка плана: {PLAN_DIR}")

    plan_file = find_plan_file()
    bank_files = sorted(
        p for folder in BANK_DIRS for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".xls")
        and not p.name.startswith("~$")
    )
    if not bank_files:
        raise RuntimeError("В папках bank_1 и bank_2 не найдено Excel-выписок.")

    print(f"\nНайдено файлов выписок: {len(bank_files)}")
    bank_rows = []
    bank_file_count = 0
    for bf in bank_files:
        if bf.suffix.lower() == ".xls":
            print(f"[ПРОПУСК] Старый формат .xls не поддерживается openpyxl: {bf.name}")
            continue
        try:
            rows = load_bank_rows(bf)
            bank_rows.extend(rows)
            bank_file_count += 1
            print(f"  Прочитан файл: {bf.name} — подходящих платежей: {len(rows)}")
        except Exception as exc:
            print(f"[ОШИБКА ЧТЕНИЯ] {bf.name}: {exc}")

    keep_vba = plan_file.suffix.lower() == ".xlsm"
    wb = openpyxl.load_workbook(plan_file, data_only=False, keep_vba=keep_vba)
    try:
        if PLAN_SHEET not in wb.sheetnames:
            raise RuntimeError(f"В книге нет листа '{PLAN_SHEET}'. Листы: {wb.sheetnames}")
        data_ws = wb[PLAN_SHEET]
        header_row = find_header_row(data_ws, ["Юрлицо", "Дата оплаты по плану до", "Долг по отгр", "рс"], 15)
        if header_row is None:
            raise RuntimeError("На листе data не найдены заголовки Юрлицо, Дата оплаты по плану до, Долг по отгр, рс.")
        columns = header_columns(data_ws, header_row)
        c_entity = find_col(columns, "Юрлицо")
        c_date = find_col(columns, "Дата оплаты по плану до")
        c_debt = find_col(columns, "Долг по отгр")
        c_rs = find_col(columns, "рс")
        if not all((c_entity, c_date, c_debt, c_rs)):
            raise RuntimeError("Не удалось определить все необходимые столбцы на листе data.")

        # Приводим отображаемое значение города к единому виду на листе data.
        # Это не меняет ключевые реквизиты и сумму, только очищает текст поля рс.
        for r in range(header_row + 1, data_ws.max_row + 1):
            cell = data_ws.cell(r, c_rs)
            if not isinstance(cell, MergedCell) and cell.value is not None:
                cleaned_city = city_from_rs(cell.value)
                if cleaned_city:
                    cell.value = cleaned_city

        groups = defaultdict(list)
        for r in range(header_row + 1, data_ws.max_row + 1):
            entity = cell_value(data_ws, r, c_entity)
            planned_date = date_key(cell_value(data_ws, r, c_date))
            rs = cell_value(data_ws, r, c_rs)
            debt = money(cell_value(data_ws, r, c_debt))
            if entity is None or planned_date is None or not rs or debt is None:
                continue
            if not any(same_entity(entity, allowed) for allowed in LEGAL_ENTITY_FILTERS):
                continue
            # Группировка строго по трём значениям: Юрлицо, дата, рс.
            groups[(str(entity).strip(), planned_date, str(rs).strip())].append((r, debt))

        print(f"\nПлановых групп для сверки: {len(groups)}")
        print(f"Банковских платежей, прошедших фильтры: {len(bank_rows)}")

        if RESULT_SHEET in wb.sheetnames:
            send_ws = wb[RESULT_SHEET]
        else:
            send_ws = wb.create_sheet(RESULT_SHEET)
            copy_sheet_layout(data_ws, send_ws, data_ws.max_column)
            # Копируем исходную строку заголовков на ту же строку.
            for rr in range(1, header_row + 1):
                if rr == header_row:
                    copy_row_values_and_style(data_ws, rr, send_ws, rr, data_ws.max_column)
            send_ws.freeze_panes = data_ws.freeze_panes

        # Оставляем на листе send только одну строку шапки; удаляем повторные копии.
        expected_header_values = [norm(data_ws.cell(header_row, c).value) for c in range(1, data_ws.max_column + 1)]
        first_send_header = find_header_row(send_ws, ["Юрлицо", "Дата оплаты по плану до", "Долг по отгр", "рс"], 50)
        if first_send_header:
            for rr in range(send_ws.max_row, first_send_header, -1):
                vals = [norm(send_ws.cell(rr, c).value) for c in range(1, data_ws.max_column + 1)]
                if vals == expected_header_values:
                    send_ws.delete_rows(rr, 1)

        # Ключи групп, уже присутствующих на листе send, чтобы повторный запуск
        # не добавлял те же группы повторно.
        existing_keys = set()
        send_header = find_header_row(send_ws, ["Юрлицо", "Дата оплаты по плану до", "Долг по отгр", "рс"], 20)
        if send_header:
            scols = header_columns(send_ws, send_header)
            se = find_col(scols, "Юрлицо")
            sd = find_col(scols, "Дата оплаты по плану до")
            sr = find_col(scols, "рс")
            # Очищаем также ранее перенесённые строки на send, чтобы старые
            # значения вида "ОБЩИЙ, МОСКВА" отображались как "МОСКВА".
            # Группа распознаётся по каждой строке данных на send.
            for rr in range(send_header + 1, send_ws.max_row + 1):
                e = cell_value(send_ws, rr, se)
                d = date_key(cell_value(send_ws, rr, sd))
                rs_cell = send_ws.cell(rr, sr) if sr else None
                rs = rs_cell.value if rs_cell is not None else None
                if rs_cell is not None and not isinstance(rs_cell, MergedCell) and rs:
                    cleaned_city = city_from_rs(rs)
                    if cleaned_city:
                        rs_cell.value = cleaned_city
                        rs = cleaned_city
                if e is not None and d is not None and rs:
                    existing_keys.add((str(e).strip(), d, str(rs).strip()))

        appended_groups = 0
        total_plan_rows = 0
        rows_to_delete = set()
        no_match_count = 0
        ambiguous_count = 0
        already_on_send_count = 0
        for key, plan_rows in groups.items():
            entity, planned_date, rs = key
            expected = sum((amount for _, amount in plan_rows), Decimal("0.00")).quantize(Decimal("0.01"))
            if key in existing_keys:
                # Группа уже перенесена на send при предыдущем запуске.
                # Если её исходные строки всё ещё есть на data, удаляем их,
                # но повторно на send не копируем.
                already_on_send_count += 1
                rows_to_delete.update(src_row for src_row, _ in plan_rows)
                print(
                    f"[УЖЕ НА SEND — УДАЛЯЮ С DATA] {entity} | "
                    f"город {city_from_rs(rs)} | сумма {expected:.2f} | "
                    f"дата плана {planned_date:%d.%m.%Y} | "
                    f"строк data к удалению: {len(plan_rows)}"
                )
                continue

            # Дата операции может отличаться от плановой даты.
            # Обязательные условия: точное совпадение суммы и города в назначении платежа.
            candidates = [b for b in bank_rows
                          if same_entity(b.get("entity"), entity)
                          and city_matches(rs, b["purpose"])
                          and (b["debit"] - expected).copy_abs() <= AMOUNT_TOLERANCE]
            # Не суммируем банковские строки: требуется одна строка с совпавшим Дебетом.
            if not candidates:
                no_match_count += 1
                continue
            if len(candidates) > 1:
                # Не выбираем произвольно, чтобы не допустить двойного переноса.
                ambiguous_count += 1
                continue

            bank = candidates[0]
            # Переносим все исходные строки группы, сохраняя значения и оформление.
            for src_row, _ in plan_rows:
                dst_row = send_ws.max_row + 1
                copy_row_values_and_style(data_ws, src_row, send_ws, dst_row, data_ws.max_column)
                rows_to_delete.add(src_row)
                total_plan_rows += 1
            # Итоговая строка: без расчётов в H и N; подпись объединена в A:J.
            total_row = send_ws.max_row + 1
            for col in range(1, data_ws.max_column + 1):
                src = data_ws.cell(header_row, col)
                dst = send_ws.cell(total_row, col)
                if src.has_style:
                    dst._style = copy(src._style)
            # Не допускаем повторного объединения диапазона при повторных запусках.
            for merged_range in list(send_ws.merged_cells.ranges):
                if merged_range.min_row == total_row and merged_range.max_row == total_row and not (merged_range.max_col < 1 or merged_range.min_col > 10):
                    send_ws.unmerge_cells(str(merged_range))
            send_ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=10)
            label_cell = send_ws.cell(total_row, 1)
            label_cell.value = "ИТОГО по группе"

            alignment = copy(label_cell.alignment)
            alignment.horizontal = "right"
            label_cell.alignment = alignment

            # Сумма группы не записывается в H или N. Если колонка долга вне A:J и не N,
            # показываем её в исходной колонке долга.
            if c_debt not in (8, 14) and c_debt > 10:
                send_ws.cell(total_row, c_debt).value = float(expected)
                send_ws.cell(total_row, c_debt).number_format = data_ws.cell(header_row + 1, c_debt).number_format if data_ws.max_row > header_row else '#,##0.00'
            # Информация о найденной банковской строке — в журнал, не в исходный лист data.
            print(f"[СОВПАДЕНИЕ] {entity} | город {city_from_rs(rs)} | сумма {expected:.2f} | "
                  f"дата плана {planned_date:%d.%m.%Y} | дата банка {bank['date']:%d.%m.%Y} | "
                  f"{bank['folder']}\\{bank['file']} лист {bank['sheet']} строка {bank['row']}")
            existing_keys.add(key)
            appended_groups += 1

        # Удаляем из data только строки тех групп, которые успешно перенесены на send.
        # Удаление снизу вверх сохраняет корректность исходных номеров строк.
        for row_idx in sorted(rows_to_delete, reverse=True):
            data_ws.delete_rows(row_idx, 1)

        # Пересчитываем L1 СРАЗУ после физического удаления строк с листа data.
        # Суммируем только числовые значения столбца L ниже строки заголовков,
        # чтобы старое значение L1 не оставалось после удаления строк.
        l1_total = Decimal("0.00")
        for rr in range(header_row + 1, data_ws.max_row + 1):
            amount = money(data_ws.cell(rr, 12).value)
            if amount is not None:
                l1_total += amount
        data_ws["L1"] = float(l1_total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
        data_ws["L1"].number_format = '#,##0.00;[Red]-#,##0.00'
        print(f"Пересчитана ячейка data!L1 после удаления строк: {l1_total:,.2f}".replace(",", " "))

        report_date = date_from_filename(plan_file)
        total_city_count, total_date_col = update_total_sheet(
            wb, data_ws, c_rs, c_debt, header_row, report_date
        )
        print(f"Лист '{TOTAL_SHEET}' обновлён: городов {total_city_count}, дата {report_date:%d.%m.%Y}, столбец {total_date_col}.")

        backup = plan_file.with_name(plan_file.stem + "_backup_before_reconcile" + plan_file.suffix)
        if not backup.exists():
            shutil.copy2(plan_file, backup)
        # Просим Excel пересчитать формулы, в том числе L1 на листе data, при открытии книги.
        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.calculation.calcMode = "auto"
        except Exception:
            pass
        # ВАЖНО: openpyxl используется только для подготовки промежуточной копии.
        # Исходный XLSM сохраняет сам Excel через COM, чтобы не перезаписывать VBA-проект
        # и прочие элементы книги средствами openpyxl.
        temp_file = plan_file.with_name(plan_file.stem + "_TEMP_FOR_EXCEL_COM" + plan_file.suffix)
        try:
            wb.save(temp_file)
        finally:
            wb.close()

        excel = None
        source_wb = None
        result_wb = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            print(f"Проверяю исходную книгу: {plan_file.resolve()}")
            source_wb = excel.Workbooks.Open(
                str(plan_file.resolve()),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True
            )
            print("Исходная книга открыта успешно.")

            print(f"Проверяю временную книгу: {temp_file.resolve()}")
            result_wb = excel.Workbooks.Open(
                str(temp_file.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True
            )
            print("Временная книга открыта успешно.")

            # Не копируем и не переименовываем treemap: это отдельная визуализация.
            # Работаем только с листами data, send и total по их точным именам.

            # Копируем data и send по содержимому. total обрабатываем отдельно ниже.
            for sheet_name in (PLAN_SHEET, RESULT_SHEET):
                try:
                    source_ws = source_wb.Worksheets(sheet_name)
                except Exception:
                    source_ws = None
                try:
                    result_ws = result_wb.Worksheets(sheet_name)
                except Exception:
                    result_ws = None
                if result_ws is None:
                    continue
                if source_ws is None:
                    result_ws.Copy(After=source_wb.Worksheets(source_wb.Worksheets.Count))
                    # Excel делает новый скопированный лист активным; используем его,
                    # а не предполагаем, что это обязательно последний по индексу.
                    excel.ActiveSheet.Name = sheet_name
                    continue

                try:
                    source_ws.UsedRange.UnMerge()
                except Exception:
                    pass
                source_ws.UsedRange.Clear()
                result_ws.UsedRange.Copy(Destination=source_ws.Range("A1"))
                source_ws.Calculate()

            # Надёжно создаём отдельный лист total непосредственно в ИСХОДНОЙ книге.
            # Не используем Worksheet.Copy как способ создания: в некоторых сценариях
            # Excel COM активным остаётся другой лист, из-за чего переименовывался не тот лист.
            try:
                result_total = result_wb.Worksheets(TOTAL_SHEET)
            except Exception:
                result_total = None
            if result_total is None:
                raise RuntimeError("Во временной книге не найден подготовленный лист 'total'.")

            # Удаляем только старый лист с точным именем total. Лист treemap не трогаем.
            try:
                existing_total = source_wb.Worksheets(TOTAL_SHEET)
            except Exception:
                existing_total = None
            if existing_total is not None:
                existing_total.Delete()

            # Явно добавляем новый лист и присваиваем имя именно объекту, который создали.
            new_total = source_wb.Worksheets.Add(
                After=source_wb.Worksheets(source_wb.Worksheets.Count)
            )
            new_total.Name = TOTAL_SHEET

            # Переносим значения и оформление расчётов с подготовленного листа.
            result_used = result_total.UsedRange
            result_used.Copy(Destination=new_total.Range("A1"))

            # После копирования листа из временной книги Excel может преобразовать
            # формулы в ссылки на TEMP_FOR_EXCEL_COM.xlsm. Перезаписываем формулы
            # текущей даты в ИСХОДНОЙ книге, чтобы они ссылались на локальный лист data.
            # Так ручное удаление строк на data будет пересчитывать total без внешних ссылок.
            try:
                from openpyxl.utils import get_column_letter
                current_total_col = None
                for cc in range(2, int(new_total.UsedRange.Columns.Count) + 1):
                    header_value = new_total.Cells(2, cc).Value
                    if isinstance(header_value, datetime):
                        if header_value.date() == report_date:
                            current_total_col = cc
                            break
                    elif header_value is not None:
                        try:
                            if parse_date(header_value).date() == report_date:
                                current_total_col = cc
                                break
                        except Exception:
                            pass
                if current_total_col is not None:
                    rs_letter = get_column_letter(c_rs)
                    debt_letter = get_column_letter(c_debt)
                    last_city_row = int(new_total.Cells(new_total.Rows.Count, 1).End(-4162).Row)  # xlUp
                    for rr in range(3, last_city_row + 1):
                        city_value = new_total.Cells(rr, 1).Value
                        if city_value is None or str(city_value).strip().upper() == "ИТОГО":
                            continue
                        # Используем ограниченные диапазоны и локальную ссылку на data.
                        new_total.Cells(rr, current_total_col).Formula = (
                            f'=ROUND(SUMIF(data!${rs_letter}$2:${rs_letter}$10000,'
                            f'"*"&$A{rr}&"*",'
                            f'data!${debt_letter}$2:${debt_letter}$10000)/1000,0)'
                        )
            except Exception as formula_error:
                print(f"Предупреждение: не удалось переписать формулы total на локальные ссылки: {formula_error}")

            try:
                new_total.Columns.AutoFit()
            except Exception:
                pass

            # Проверяем, что лист действительно присутствует в исходной книге перед сохранением.
            source_wb.Worksheets(TOTAL_SHEET)
            print("Создан и проверен отдельный лист 'total' с расчётами по городам.")

            # Пересчёт выполняем через приложение Excel: свойства
            # ForceFullCalculation и FullCalculationOnLoad не поддерживаются
            # одинаково всеми версиями Excel/COM для объекта Workbook.
            try:
                excel.CalculateFullRebuild()
            except Exception as calc_error:
                print(f"Предупреждение: полный пересчёт Excel не выполнен: {calc_error}")
            source_wb.Save()
        finally:
            if result_wb is not None:
                result_wb.Close(SaveChanges=False)
            if source_wb is not None:
                source_wb.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
            try:
                if temp_file.exists():
                    temp_file.unlink()
            except OSError:
                print(f"Предупреждение: временный файл не удалён: {temp_file}")

        print("\n" + "-" * 68)
        print("ИТОГ ОБРАБОТКИ")
        print(f"Файлов выписок прочитано: {bank_file_count} из {len(bank_files)}")
        print(f"Банковских платежей отобрано: {len(bank_rows)}")
        print(f"Плановых групп: {len(groups)}")
        print(f"Найдено совпадений и перенесено групп: {appended_groups}")
        print(f"Перенесено строк на '{RESULT_SHEET}': {total_plan_rows}")
        print(f"Физически удалено строк с '{PLAN_SHEET}': {len(rows_to_delete)}")
        print(f"Групп без совпадения по сумме и городу: {no_match_count}")
        print(f"Групп с несколькими подходящими платежами (пропущены): {ambiguous_count}")
        print(f"Групп, уже находившихся на '{RESULT_SHEET}': {already_on_send_count}")
        print(f"Файл сохранён: {plan_file}")
        print("=" * 68)
    finally:
        wb.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nОШИБКА: {exc}")
    input("\nНажмите Enter для выхода...")
