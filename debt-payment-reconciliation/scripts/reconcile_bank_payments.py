
# -*- coding: utf-8 -*-
"""Сверка плана оплат с банковскими выписками.

Группа переносится на лист send при совпадении суммы и города в назначении платежа;
дата операции банка может отличаться от плановой даты. Перенесённые строки
физически удаляются с листа data. Перед сохранением создаётся резервная копия.
"""
from __future__ import annotations

import os
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
# Это синтетические значения для демонстрационного набора. Не добавляйте
# реальные реквизиты в публичный код.
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
    text = re.sub(r'["«»„“”]', "", text)
    text = re.sub(r"\b(ооо|ао|пао|зао|оао)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def same_entity(left, right) -> bool:
    """Сопоставляет юрлица по точному нормализованному названию."""
    a, b = entity_key(left), entity_key(right)
    return bool(a and b and a == b)


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


CITY_ALIASES = {
    "\u041c\u0421\u041a": "\u041c\u041e\u0421\u041a\u0412\u0410",
    "\u0415\u041a\u0411": "\u0415\u041a\u0410\u0422\u0415\u0420\u0418\u041d\u0411\u0423\u0420\u0413",
    "\u0421\u041f\u0411": "\u0421\u0410\u041d\u041a\u0422-\u041f\u0415\u0422\u0415\u0420\u0411\u0423\u0420\u0413",
    "МОСКВА-СОСЕНКИ": "МОСКВА",
}


def normalize_city_aliases(value):
    text = str(value or "").replace("\u00a0", " ")
    for alias, full_name in CITY_ALIASES.items():
        text = re.sub(
            rf"(?<!\w){re.escape(alias)}(?!\w)",
            full_name,
            text,
            flags=re.IGNORECASE,
        )
    return re.sub(r"\s+", " ", text).strip()


def city_from_rs(value):
    """Очищает служебный префикс и приводит город к единому названию."""
    city = str(value or "").replace("\u00a0", " ").strip()
    city = re.sub(
        r"^\s*ОБЩИЙ\s*[,;:-]\s*",
        "",
        city,
        flags=re.IGNORECASE,
    )
    city = normalize_city_aliases(city)
    return city.strip(" ,;:-").upper()


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
    raise RuntimeError(
        f"Не удалось определить дату из названия файла: {path.name}. "
        "Используй дату в формате ДД.ММ.ГГ или ДД.ММ.ГГГГ."
    )


def update_total_sheet(wb, data_ws, c_rs, c_debt, header_row, report_date, cached_total_ws=None):
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
            if city and city.upper() not in {
                "ГОРОД",
                "ИТОГО",
                "ДАННЫЕ ПО ЗАДОЛЖЕННОСТИ ПО ГОРОДАМ, ТЫС.РУБ.",
            }:
                old_cities[r] = city

        # Поддерживаем оба заголовка, которые использовались в версиях скрипта.
        # Сравнение через norm учитывает регистр и лишние пробелы.
        known_total_titles = {
            norm("ДАННЫЕ ПО ГОРОДАМ, тыс.руб."),
            norm("ДАННЫЕ ПО ЗАДОЛЖЕННОСТИ ПО ГОРОДАМ, тыс.руб."),
        }
        title_is_known = norm(ws.cell(1, 1).value) in known_total_titles
        looks_like_total = title_is_known and bool(old_dates) and bool(old_cities)

        # Если на листе уже есть даты и города, но заголовок не распознан,
        # не очищаем его автоматически: это может быть чужая/ручная история.
        if old_dates and old_cities and not title_is_known:
            raise RuntimeError(
                "Лист 'total' содержит даты и города, но его заголовок A1 "
                "не распознан. Лист не изменён, чтобы не потерять историю. "
                f"Текущее значение A1: {ws.cell(1, 1).value!r}"
            )

        history_cols = []
        if looks_like_total:
            if cached_total_ws is None:
                raise RuntimeError(
                    "Для сохранения истории листа 'total' не загружены "
                    "кэшированные значения книги (data_only=True)."
                )

            for c, d in old_dates.items():
                values = {}
                for r, city in old_cities.items():
                    formula_or_value = ws.cell(r, c).value
                    cached_value = cached_total_ws.cell(r, c).value

                    if isinstance(formula_or_value, str) and formula_or_value.startswith("="):
                        if not isinstance(cached_value, (int, float, Decimal)):
                            raise RuntimeError(
                                "Не удалось получить сохранённый числовой результат "
                                "старой формулы на листе 'total': "
                                f"ячейка {ws.cell(r, c).coordinate}, город {city!r}, "
                                f"дата {d:%d.%m.%Y}. Открой книгу в Excel, дождись "
                                "пересчёта формул, сохрани её и повтори запуск. "
                                "История не будет очищена."
                            )
                        values[city.casefold()] = cached_value
                    elif isinstance(cached_value, (int, float, Decimal)):
                        values[city.casefold()] = cached_value
                    # Пустые ячейки остаются пустыми: город мог появиться позднее.

                history_cols.append((d, values))
        else:
            # Пустой лист total или лист без исторических дат/городов:
            # начинаем формировать структуру заново.
            history_cols = []
            old_cities = {}

        for merged_range in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged_range))
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)
        for col in list(ws.column_dimensions):
            del ws.column_dimensions[col]

    totals = defaultdict(lambda: Decimal("0.00"))
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
    matching_index = next(
        (i for i, (dt, _) in enumerate(date_columns) if dt == report_date),
        None,
    )
    if matching_index is None:
        date_columns.append((report_date, {}))
        current_index = len(date_columns) - 1
    else:
        current_index = matching_index
        date_columns[current_index] = (report_date, {})

    priority = {"МОСКВА": 0, "САНКТ-ПЕТЕРБУРГ": 1}
    cities_sorted = sorted(
        all_cities,
        key=lambda x: (priority.get(x.upper(), 2), x.casefold()),
    )

    ws.cell(1, 1).value = "ДАННЫЕ ПО ЗАДОЛЖЕННОСТИ ПО ГОРОДАМ, тыс.руб."
    ws.cell(2, 1).value = "Город"
    for col, (dt, _) in enumerate(date_columns, start=2):
        ws.cell(2, col).value = dt
        ws.cell(2, col).number_format = "dd.mm.yyyy"

    for row, city in enumerate(cities_sorted, start=3):
        ws.cell(row, 1).value = city
        for col, (_, values) in enumerate(date_columns, start=2):
            if col - 2 == current_index:
                # Формула Excel пересчитывается при ручном удалении/изменении строк на data.
                # Значение города на листе data предварительно нормализовано.
                rs_letter = get_column_letter(c_rs)
                debt_letter = get_column_letter(c_debt)
                city_cell = f"$A{row}"

                # Используем точное совпадение, а не поиск по фрагменту названия.
                # Диапазон охватывает все строки листа Excel.
                rs_range = f"data!${rs_letter}$2:${rs_letter}$1048576"
                debt_range = f"data!${debt_letter}$2:${debt_letter}$1048576"
                ws.cell(row, col).value = (
                    f"=ROUND(SUMIF({rs_range},{city_cell},{debt_range})/1000,0)"
                )
            elif city.casefold() in values:
                ws.cell(row, col).value = values[city.casefold()]

    # Строка ИТОГО всегда идёт сразу под последним городом.
    total_row = len(cities_sorted) + 3
    ws.cell(total_row, 1).value = "ИТОГО"
    for col in range(2, len(date_columns) + 2):
        letter = get_column_letter(col)
        ws.cell(total_row, col).value = (
            f"=SUM({letter}3:{letter}{total_row - 1})"
        )

    from openpyxl.styles import Border, Side
    for row in ws.iter_rows(
        min_row=1,
        max_row=max(2, total_row),
        min_col=1,
        max_col=max(2, len(date_columns) + 1),
    ):
        for cell in row:
            cell.font = Font(name="Times New Roman", size=12, bold=True)
            cell.alignment = Alignment(vertical="center")
            if cell.row >= 3 and cell.column >= 2:
                cell.number_format = "# ##0;[Red]-# ##0"

    total_border = Border(top=Side(style="thin"))
    for col in range(1, len(date_columns) + 2):
        total_cell = ws.cell(total_row, col)
        total_cell.font = Font(name="Times New Roman", size=12, bold=True)
        total_cell.border = total_border
        if col >= 2:
            total_cell.number_format = "# ##0;[Red]-# ##0"

    ws.column_dimensions["A"].width = max(
        18,
        min(
            40,
            max([len(str(c)) for c in cities_sorted] + [5, len("ИТОГО")]) + 3,
        ),
    )
    for col in range(2, len(date_columns) + 2):
        ws.column_dimensions[get_column_letter(col)].width = 16

    ws.freeze_panes = "B3"
    ws.sheet_view.showGridLines = False
    return len(totals), len(date_columns) + 1


def city_token_key(token):
    """Убирает типовые окончания падежа, сохраняя основу названия города."""
    suffixes = (
        "иями", "ями", "ами", "ого", "ему", "ому", "ыми", "ими",
        "ах", "ях", "ам", "ям", "ов", "ев", "ом", "ем",
        "ой", "ей", "ью", "ия", "ие", "а", "я", "у", "ю", "е", "ы", "и", "ь",
    )
    for suffix in suffixes:
        if token.endswith(suffix) and len(token) - len(suffix) >= 3:
            return token[:-len(suffix)]
    return token


def city_matches(rs_value, payment_purpose):
    """Ищет полное название города в назначении с учётом падежных окончаний."""
    city_text = city_from_rs(rs_value)
    purpose_text = normalize_city_aliases(payment_purpose)

    city_tokens = [
        city_token_key(token)
        for token in re.findall(
            r"[^\W_]+",
            city_text.casefold(),
            flags=re.UNICODE,
        )
    ]
    purpose_tokens = [
        city_token_key(token)
        for token in re.findall(
            r"[^\W_]+",
            purpose_text.casefold(),
            flags=re.UNICODE,
        )
    ]

    if (
        not city_tokens
        or not purpose_tokens
        or len(city_tokens) > len(purpose_tokens)
    ):
        return False

    width = len(city_tokens)
    return any(
        purpose_tokens[index:index + width] == city_tokens
        for index in range(len(purpose_tokens) - width + 1)
    )


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
    candidates = [
        p for p in PLAN_DIR.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xlsm")
        and not p.name.startswith("~$")
        and "_backup_before_reconcile" not in p.stem.lower()
        and "_temp_for_excel_com" not in p.stem.lower()
    ]
    if len(candidates) != 1:
        raise RuntimeError(
            f"В папке {PLAN_DIR} должен быть ровно один файл Excel; найдено: {len(candidates)}."
        )
    return candidates[0]


def load_bank_rows(bank_file):
    """Читает два варианта структуры выписок: многострочную и обычную шапку."""
    keep_vba = bank_file.suffix.lower() == ".xlsm"
    wb = openpyxl.load_workbook(
        bank_file,
        data_only=True,
        read_only=False,
        keep_vba=keep_vba,
    )
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
                [
                    "Дата операции/Posting date",
                    "Дебет / Debit",
                    "Реквизиты корреспондента / Counter party details",
                    "Основание операции (назначение платежа) / Payment details",
                ],
                BANK_HEADER_SCAN_ROWS,
            )

            # Универсальное распознавание обычной и многострочной шапки.
            if header_row is None:
                for r in range(1, min(ws.max_row, BANK_HEADER_SCAN_ROWS) + 1):
                    row_text = " | ".join(
                        norm(ws.cell(r, c).value)
                        for c in range(1, ws.max_column + 1)
                    )
                    has_debit = "дебет" in row_text
                    has_purpose = any(
                        term in row_text
                        for term in (
                            "назначение",
                            "основание операции",
                            "payment details",
                        )
                    )
                    has_date = "дата" in row_text or "posting date" in row_text
                    has_inn = (
                        "инн" in row_text
                        or "counter party details" in row_text
                        or "реквизиты корреспондента" in row_text
                    )
                    if has_debit and has_purpose and (has_date or has_inn):
                        header_row = r
                        break

            if header_row is None:
                # Запасной вариант для выписки с нумерованными колонками.
                for r in range(1, min(ws.max_row, BANK_HEADER_SCAN_ROWS) + 1):
                    vals = [
                        norm(ws.cell(r, c).value)
                        for c in range(1, ws.max_column + 1)
                    ]
                    if (
                        any(v in ("9", "дебет", "дебет / debit") for v in vals)
                        and ws.max_column >= 23
                    ):
                        header_row = r
                        break

            if header_row is None:
                print(
                    f"  [ПРОПУСК] Не удалось определить заголовки: "
                    f"{bank_file.name}, лист {ws.title}"
                )
                continue

            col_date = col_debit = col_inn = col_purpose = None
            for r in range(
                max(1, header_row - 3),
                min(ws.max_row, header_row + 1) + 1,
            ):
                for c in range(1, ws.max_column + 1):
                    value = norm(ws.cell(r, c).value)

                    if (
                        value in ("дата", "дата операции", "posting date")
                        or "дата операции" in value
                        or "posting date" in value
                    ):
                        col_date = c

                    if (
                        value in ("дебет / debit", "дебет")
                        or ("дебет" in value and col_debit is None)
                    ):
                        col_debit = c

                    if (
                        "инн контрагента" in value
                        or "реквизиты корреспондента" in value
                        or "counter party details" in value
                    ):
                        col_inn = c

                    if (
                        value == "назначение"
                        or "назначение платежа" in value
                        or "основание операции" in value
                        or "payment details" in value
                    ):
                        col_purpose = c

            # Фиксированные позиции — только резерв для прежнего формата выписки.
            if col_date is None and ws.max_column >= 23:
                col_date = 3
            if col_inn is None and ws.max_column >= 23:
                col_inn = 13
            if col_debit is None and ws.max_column >= 23:
                col_debit = 17
            if col_purpose is None and ws.max_column >= 23:
                col_purpose = 23

            if not all((col_date, col_debit, col_inn, col_purpose)):
                print(
                    f"  [ПРОПУСК] Не найдены нужные столбцы: "
                    f"{bank_file.name}, лист {ws.title}"
                )
                continue

            for r in range(header_row + 1, ws.max_row + 1):
                operation_date = date_key(cell_value(ws, r, col_date))
                debit = money(cell_value(ws, r, col_debit))
                inn_text = str(cell_value(ws, r, col_inn) or "")
                purpose = cell_value(ws, r, col_purpose)

                if operation_date is None or debit is None or not purpose:
                    continue

                # ИНН ищется как отдельная последовательность цифр, а не подстрока.
                if not any(
                    re.search(rf"(?<!\d){re.escape(inn)}(?!\d)", inn_text)
                    for inn in COUNTERPARTY_INNS
                ):
                    continue

                found.append({
                    "file": bank_file.name,
                    "sheet": ws.title,
                    "row": r,
                    "date": operation_date,
                    "debit": debit,
                    "inn_text": inn_text,
                    "purpose": str(purpose),
                    "entity": bank_entity,
                    "folder": bank_file.parent.name,
                })
    finally:
        wb.close()

    return found


def create_backup(plan_file: Path) -> Path:
    """Создаёт отдельную резервную копию перед каждой обработкой, не перезаписывая старую."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_stem = f"{plan_file.stem}_backup_before_reconcile_{timestamp}"
    backup = plan_file.with_name(base_stem + plan_file.suffix)
    counter = 1

    while backup.exists():
        backup = plan_file.with_name(
            f"{base_stem}_{counter}{plan_file.suffix}"
        )
        counter += 1

    shutil.copy2(plan_file, backup)
    return backup


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
        raise RuntimeError(
            "Не найдены папки банковских выписок:\n"
            + "\n".join(map(str, missing_dirs))
        )
    if not PLAN_DIR.exists():
        raise RuntimeError(f"Не найдена папка плана: {PLAN_DIR}")

    plan_file = find_plan_file()
    bank_files = sorted(
        p for folder in BANK_DIRS for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in (".xlsx", ".xlsm", ".xls")
        and not p.name.startswith("~$")
    )
    if not bank_files:
        raise RuntimeError("В папках bank_1 и bank_2 не найдено Excel-выписок.")

    print(f"\nНайдено файлов выписок: {len(bank_files)}")
    bank_rows = []
    bank_file_count = 0
    bank_read_errors = []

    for bank_file in bank_files:
        if bank_file.suffix.lower() == ".xls":
            bank_read_errors.append(
                f"{bank_file.name}: формат .xls не поддерживается openpyxl"
            )
            continue

        try:
            rows = load_bank_rows(bank_file)
            bank_rows.extend(rows)
            bank_file_count += 1
            print(
                f"  Прочитан файл: {bank_file.name} — "
                f"подходящих платежей: {len(rows)}"
            )
        except Exception as exc:
            bank_read_errors.append(f"{bank_file.name}: {exc}")
            print(f"[ОШИБКА ЧТЕНИЯ] {bank_file.name}: {exc}")

    # Нельзя продолжать сверку по неполному набору выписок: это может привести
    # к ложному выводу об отсутствии платежа и неправильному переносу строк.
    if bank_read_errors:
        raise RuntimeError(
            "Не все банковские выписки удалось прочитать. "
            "Сверка остановлена без изменения плана:\n- "
            + "\n- ".join(bank_read_errors)
        )
    if bank_file_count == 0:
        raise RuntimeError("Не удалось прочитать ни одной банковской выписки.")
    if not bank_rows:
        raise RuntimeError(
            "В прочитанных выписках не найдено платежей с заданными ИНН. "
            "Сверка остановлена, чтобы не считать все плановые группы неоплаченными."
        )

    # Резервируем исходную книгу до любых изменений в памяти и до открытия Excel COM.
    backup = create_backup(plan_file)
    print(f"Создана резервная копия: {backup.name}")

    keep_vba = plan_file.suffix.lower() == ".xlsm"
    wb = None
    wb_values = None
    temp_dir = None

    try:
        # Основная книга нужна с формулами; отдельная копия только для чтения
        # отдаёт сохранённые Excel результаты формул листа total.
        wb = openpyxl.load_workbook(
            plan_file,
            data_only=False,
            keep_vba=keep_vba,
        )
        wb_values = openpyxl.load_workbook(
            plan_file,
            data_only=True,
            keep_vba=keep_vba,
        )
        if PLAN_SHEET not in wb.sheetnames:
            raise RuntimeError(
                f"В книге нет листа '{PLAN_SHEET}'. Листы: {wb.sheetnames}"
            )

        data_ws = wb[PLAN_SHEET]
        header_row = find_header_row(
            data_ws,
            ["Юрлицо", "Дата оплаты по плану до", "Долг по отгр"],
            15,
        )
        if header_row is None:
            raise RuntimeError(
                "На листе data не найдены обязательные заголовки: "
                "Юрлицо, Дата оплаты по плану до, Долг по отгр."
            )

        columns = header_columns(data_ws, header_row)
        c_entity = find_col(columns, "Юрлицо")
        c_date = find_col(columns, "Дата оплаты по плану до")
        c_debt = find_col(columns, "Долг по отгр")
        c_rs = find_col(columns, "Город", "рс")
        if not all((c_entity, c_date, c_debt, c_rs)):
            raise RuntimeError(
                "Не удалось определить все необходимые столбцы на листе data."
            )

        # Приводим город к единому виду; исходные строки и суммы не меняем.
        for row in range(header_row + 1, data_ws.max_row + 1):
            cell = data_ws.cell(row, c_rs)
            if not isinstance(cell, MergedCell) and cell.value is not None:
                cleaned_city = city_from_rs(cell.value)
                if cleaned_city:
                    cell.value = cleaned_city

        # Ключ группировки нормализован: юрлицо, дата плана и город.
        groups = defaultdict(list)
        group_display_names = {}

        for row in range(header_row + 1, data_ws.max_row + 1):
            entity = cell_value(data_ws, row, c_entity)
            planned_date = date_key(cell_value(data_ws, row, c_date))
            city = city_from_rs(cell_value(data_ws, row, c_rs))
            debt = money(cell_value(data_ws, row, c_debt))

            normalized_entity = entity_key(entity)
            if (
                not normalized_entity
                or planned_date is None
                or not city
                or debt is None
            ):
                continue

            if not any(
                same_entity(normalized_entity, allowed)
                for allowed in LEGAL_ENTITY_FILTERS
            ):
                continue

            key = (normalized_entity, planned_date, city)
            groups[key].append((row, debt))
            group_display_names.setdefault(key, str(entity).strip())

        print(f"\nПлановых групп для сверки: {len(groups)}")
        print(f"Банковских платежей, прошедших фильтры: {len(bank_rows)}")

        # Лист send создаётся в промежуточной книге, если его ещё нет.
        if RESULT_SHEET in wb.sheetnames:
            send_ws = wb[RESULT_SHEET]
        else:
            send_ws = wb.create_sheet(RESULT_SHEET)
            copy_sheet_layout(data_ws, send_ws, data_ws.max_column)
            copy_row_values_and_style(
                data_ws,
                header_row,
                send_ws,
                header_row,
                data_ws.max_column,
            )
            send_ws.freeze_panes = data_ws.freeze_panes

        expected_header_values = [
            norm(data_ws.cell(header_row, col).value)
            for col in range(1, data_ws.max_column + 1)
        ]
        first_send_header = find_header_row(
            send_ws,
            ["Юрлицо", "Дата оплаты по плану до", "Долг по отгр"],
            50,
        )
        if first_send_header is None:
            raise RuntimeError(
                f"На листе '{RESULT_SHEET}' не найдена строка заголовков."
            )

        # Удаляем только повторные копии строки заголовков.
        for row in range(send_ws.max_row, first_send_header, -1):
            values = [
                norm(send_ws.cell(row, col).value)
                for col in range(1, data_ws.max_column + 1)
            ]
            if values == expected_header_values:
                send_ws.delete_rows(row, 1)

        send_header = find_header_row(
            send_ws,
            ["Юрлицо", "Дата оплаты по плану до", "Долг по отгр"],
            50,
        )
        if send_header is None:
            raise RuntimeError(
                f"После подготовки листа '{RESULT_SHEET}' не найдена шапка."
            )

        send_columns = header_columns(send_ws, send_header)
        send_entity_col = find_col(send_columns, "Юрлицо")
        send_date_col = find_col(send_columns, "Дата оплаты по плану до")
        send_debt_col = find_col(send_columns, "Долг по отгр")
        send_city_col = find_col(send_columns, "Город", "рс")
        if not all((
            send_entity_col,
            send_date_col,
            send_debt_col,
            send_city_col,
        )):
            raise RuntimeError(
                f"На листе '{RESULT_SHEET}' отсутствуют нужные столбцы."
            )

        # Сохраняем суммы уже перенесённых групп. Одного совпадения ключа
        # недостаточно: сумма тоже должна совпадать, иначе это конфликт данных.
        existing_send_keys = set()
        existing_send_amounts = defaultdict(lambda: Decimal("0.00"))

        for row in range(send_header + 1, send_ws.max_row + 1):
            label = norm(cell_value(send_ws, row, 1))
            if label == norm("ИТОГО по группе"):
                continue

            entity = cell_value(send_ws, row, send_entity_col)
            planned_date = date_key(cell_value(send_ws, row, send_date_col))
            city = city_from_rs(cell_value(send_ws, row, send_city_col))
            if not entity or planned_date is None or not city:
                continue

            key = (entity_key(entity), planned_date, city)
            if not key[0]:
                continue

            existing_send_keys.add(key)
            amount = money(cell_value(send_ws, row, send_debt_col))
            if amount is not None:
                existing_send_amounts[key] += amount

            city_cell = send_ws.cell(row, send_city_col)
            if (
                not isinstance(city_cell, MergedCell)
                and city_cell.value is not None
            ):
                city_cell.value = city

        # Сначала строим соответствия для всех групп. Это позволяет обнаружить
        # ситуацию, когда одна банковская строка подходит сразу нескольким группам.
        candidates_by_group = {}
        bank_owners = defaultdict(list)

        for key, plan_rows in groups.items():
            if key in existing_send_keys:
                continue

            entity, planned_date, city = key
            expected = sum(
                (amount for _, amount in plan_rows),
                Decimal("0.00"),
            ).quantize(Decimal("0.01"))

            candidates = [
                bank for bank in bank_rows
                if same_entity(bank.get("entity"), entity)
                and city_matches(city, bank["purpose"])
                and (bank["debit"] - expected).copy_abs() <= AMOUNT_TOLERANCE
            ]
            candidates_by_group[key] = candidates

            # Регистрируем все возможные связи, в том числе неоднозначные.
            # Это не позволяет другой группе забрать один из таких платежей.
            for bank in candidates:
                bank_id = (
                    bank["folder"].casefold(),
                    bank["file"],
                    bank["sheet"],
                    bank["row"],
                )
                bank_owners[bank_id].append(key)

        appended_groups = 0
        total_plan_rows = 0
        rows_to_delete = set()
        no_match_count = 0
        ambiguous_count = 0
        already_on_send_count = 0
        send_conflict_count = 0
        bank_collision_count = 0

        for key, plan_rows in groups.items():
            entity_key_value, planned_date, city = key
            display_entity = group_display_names.get(key, entity_key_value)
            expected = sum(
                (amount for _, amount in plan_rows),
                Decimal("0.00"),
            ).quantize(Decimal("0.01"))

            # Повторный запуск удаляет строки с data только при совпадении
            # и ключа, и суммы с уже перенесённой группой на send.
            if key in existing_send_keys:
                saved_amount = existing_send_amounts.get(
                    key,
                    Decimal("0.00"),
                ).quantize(Decimal("0.01"))

                if (saved_amount - expected).copy_abs() <= AMOUNT_TOLERANCE:
                    already_on_send_count += 1
                    rows_to_delete.update(
                        source_row for source_row, _ in plan_rows
                    )
                    print(
                        f"[УЖЕ НА SEND — УДАЛЯЮ С DATA] {display_entity} | "
                        f"город {city} | сумма {expected:.2f} | "
                        f"дата плана {planned_date:%d.%m.%Y} | "
                        f"строк data к удалению: {len(plan_rows)}"
                    )
                else:
                    send_conflict_count += 1
                    print(
                        f"[КОНФЛИКТ С SEND — ПРОПУСК] {display_entity} | "
                        f"город {city} | сумма data {expected:.2f} | "
                        f"сумма send {saved_amount:.2f} | "
                        f"дата плана {planned_date:%d.%m.%Y}"
                    )
                continue

            candidates = candidates_by_group.get(key, [])
            if not candidates:
                no_match_count += 1
                continue

            if len(candidates) > 1:
                ambiguous_count += 1
                continue

            bank = candidates[0]
            bank_id = (
                bank["folder"].casefold(),
                bank["file"],
                bank["sheet"],
                bank["row"],
            )

            # Если один платёж подходит нескольким группам, не выбираем
            # произвольного получателя: все такие группы остаются на data.
            if len(bank_owners[bank_id]) > 1:
                ambiguous_count += 1
                bank_collision_count += 1
                continue

            # Переносим все исходные строки группы, сохраняя значения и оформление.
            for source_row, _ in plan_rows:
                destination_row = send_ws.max_row + 1
                copy_row_values_and_style(
                    data_ws,
                    source_row,
                    send_ws,
                    destination_row,
                    data_ws.max_column,
                )
                rows_to_delete.add(source_row)
                total_plan_rows += 1

            # Итоговая строка группы. Подпись занимает объединённые ячейки A:J.
            total_row = send_ws.max_row + 1
            for col in range(1, data_ws.max_column + 1):
                source_cell = data_ws.cell(header_row, col)
                target_cell = send_ws.cell(total_row, col)
                if source_cell.has_style:
                    target_cell._style = copy(source_cell._style)

            for merged_range in list(send_ws.merged_cells.ranges):
                if (
                    merged_range.min_row == total_row
                    and merged_range.max_row == total_row
                    and not (
                        merged_range.max_col < 1
                        or merged_range.min_col > 10
                    )
                ):
                    send_ws.unmerge_cells(str(merged_range))

            send_ws.merge_cells(
                start_row=total_row,
                start_column=1,
                end_row=total_row,
                end_column=10,
            )
            label_cell = send_ws.cell(total_row, 1)
            label_cell.value = "ИТОГО по группе"

            alignment = copy(label_cell.alignment)
            alignment.horizontal = "right"
            label_cell.alignment = alignment

            # В H и N сумма группы не ставится; в другой колонке долга — ставится.
            if c_debt not in (8, 14) and c_debt > 10:
                send_ws.cell(total_row, c_debt).value = float(expected)
                send_ws.cell(total_row, c_debt).number_format = (
                    data_ws.cell(header_row + 1, c_debt).number_format
                    if data_ws.max_row > header_row
                    else "#,##0.00"
                )

            print(
                f"[СОВПАДЕНИЕ] {display_entity} | город {city} | "
                f"сумма {expected:.2f} | "
                f"дата плана {planned_date:%d.%m.%Y} | "
                f"дата банка {bank['date']:%d.%m.%Y} | "
                f"{bank['folder']}\\{bank['file']} "
                f"лист {bank['sheet']} строка {bank['row']}"
            )
            appended_groups += 1

        # Удаляем только строки подтверждённых групп и только снизу вверх.
        for row_index in sorted(rows_to_delete, reverse=True):
            data_ws.delete_rows(row_index, 1)

        # Пересчитываем L1 после физического удаления строк.
        l1_total = Decimal("0.00")
        for row in range(header_row + 1, data_ws.max_row + 1):
            amount = money(data_ws.cell(row, 12).value)
            if amount is not None:
                l1_total += amount

        l1_total = l1_total.quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
        data_ws["L1"] = float(l1_total)
        data_ws["L1"].number_format = "#,##0.00;[Red]-#,##0.00"
        print(
            f"Пересчитана ячейка data!L1 после удаления строк: "
            f"{l1_total:,.2f}".replace(",", " ")
        )

        report_date = date_from_filename(plan_file)
        cached_total_ws = (
            wb_values[TOTAL_SHEET]
            if wb_values is not None and TOTAL_SHEET in wb_values.sheetnames
            else None
        )
        total_city_count, total_date_col = update_total_sheet(
            wb,
            data_ws,
            c_rs,
            c_debt,
            header_row,
            report_date,
            cached_total_ws=cached_total_ws,
        )

        # Кэш уже считан; закрываем вторую книгу до дальнейшей обработки.
        if wb_values is not None:
            wb_values.close()
            wb_values = None

        print(
            f"Лист '{TOTAL_SHEET}' обновлён: городов {total_city_count}, "
            f"дата {report_date:%d.%m.%Y}, столбец {total_date_col}."
        )

        # Временные файлы размещаются в отдельной служебной папке внутри PLAN_DIR:
        # они не будут приняты за второй исходный план при следующем запуске.
        temp_dir = Path(tempfile.mkdtemp(prefix=".reconcile_", dir=PLAN_DIR))
        temp_file = temp_dir / plan_file.name
        staging_file = temp_dir / f"staged_{plan_file.name}"

        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.calculation.calcMode = "auto"
        except Exception:
            pass

        wb.save(temp_file)
        wb.close()
        wb = None

        # Excel COM работает с копией исходной книги. Оригинал заменяется только
        # после успешного сохранения и закрытия всех COM-книг.
        shutil.copy2(plan_file, staging_file)

        excel = None
        source_wb = None
        result_wb = None
        cleanup_errors = []

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False

            print(f"Проверяю рабочую копию исходной книги: {staging_file.resolve()}")
            source_wb = excel.Workbooks.Open(
                str(staging_file.resolve()),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
            )
            print("Рабочая копия исходной книги открыта успешно.")

            print(f"Проверяю временную книгу: {temp_file.resolve()}")
            result_wb = excel.Workbooks.Open(
                str(temp_file.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )
            print("Временная книга открыта успешно.")

            # Переносим листы data и send. Если send отсутствовал в оригинале,
            # копируем его из подготовленной книги целиком.
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
                    raise RuntimeError(
                        f"Во временной книге отсутствует лист '{sheet_name}'."
                    )

                if source_ws is None:
                    result_ws.Copy(
                        After=source_wb.Worksheets(
                            source_wb.Worksheets.Count
                        )
                    )
                    excel.ActiveSheet.Name = sheet_name
                    continue

                try:
                    source_ws.UsedRange.UnMerge()
                except Exception:
                    pass

                source_ws.UsedRange.Clear()
                result_ws.UsedRange.Copy(
                    Destination=source_ws.Range("A1")
                )
                source_ws.Calculate()

            # Создаём лист total в рабочей копии; treemap не изменяем.
            try:
                result_total = result_wb.Worksheets(TOTAL_SHEET)
            except Exception:
                result_total = None

            if result_total is None:
                raise RuntimeError(
                    f"Во временной книге не найден подготовленный лист '{TOTAL_SHEET}'."
                )

            try:
                existing_total = source_wb.Worksheets(TOTAL_SHEET)
            except Exception:
                existing_total = None

            if existing_total is not None:
                existing_total.Delete()

            new_total = source_wb.Worksheets.Add(
                After=source_wb.Worksheets(source_wb.Worksheets.Count)
            )
            new_total.Name = TOTAL_SHEET

            result_used = result_total.UsedRange
            result_used.Copy(Destination=new_total.Range("A1"))

            # Перезаписываем формулы текущей даты, чтобы они ссылались на лист
            # data в рабочей книге, а не на промежуточный файл.
            current_total_col = None
            for col in range(
                2,
                int(new_total.UsedRange.Columns.Count) + 1,
            ):
                header_value = new_total.Cells(2, col).Value
                if date_key(header_value) == report_date:
                    current_total_col = col
                    break

            if current_total_col is None:
                raise RuntimeError(
                    "На листе total не найден столбец с датой текущего плана."
                )

            from openpyxl.utils import get_column_letter
            rs_letter = get_column_letter(c_rs)
            debt_letter = get_column_letter(c_debt)
            last_city_row = int(
                new_total.Cells(
                    new_total.Rows.Count,
                    1,
                ).End(-4162).Row
            )  # xlUp

            for row in range(3, last_city_row + 1):
                city_value = new_total.Cells(row, 1).Value
                if (
                    city_value is None
                    or str(city_value).strip().upper() == "ИТОГО"
                ):
                    continue

                new_total.Cells(row, current_total_col).Formula = (
                    f"=ROUND(SUMIF(data!${rs_letter}$2:${rs_letter}$1048576,"
                    f"$A{row},"
                    f"data!${debt_letter}$2:${debt_letter}$1048576)/1000,0)"
                )

            try:
                new_total.Columns.AutoFit()
            except Exception:
                pass

            source_wb.Worksheets(TOTAL_SHEET)
            print(
                f"Создан и проверен отдельный лист '{TOTAL_SHEET}'."
            )

            try:
                excel.CalculateFullRebuild()
            except Exception as calc_error:
                print(
                    f"Предупреждение: полный пересчёт Excel не выполнен: "
                    f"{calc_error}"
                )

            source_wb.Save()
            print("Рабочая копия книги сохранена.")
        finally:
            # Каждая операция очистки выполняется независимо: ошибка закрытия
            # одной книги не должна мешать закрытию другой и завершению Excel.
            for workbook, workbook_label in (
                (result_wb, "временная книга"),
                (source_wb, "рабочая копия"),
            ):
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception as close_error:
                        cleanup_errors.append(
                            f"Не удалось закрыть {workbook_label}: {close_error}"
                        )

            if excel is not None:
                try:
                    excel.Quit()
                except Exception as quit_error:
                    cleanup_errors.append(
                        f"Не удалось завершить Excel: {quit_error}"
                    )

        if cleanup_errors:
            raise RuntimeError(
                "Рабочая копия не заменяет исходный файл, поскольку Excel "
                "не удалось корректно закрыть:\n- "
                + "\n- ".join(cleanup_errors)
            )

        # Атомарная замена: если Excel не смог подготовить копию, оригинал остаётся.
        os.replace(staging_file, plan_file)
        print("Исходная книга обновлена после успешной обработки рабочей копии.")

        print("\n" + "-" * 68)
        print("ИТОГ ОБРАБОТКИ")
        print(f"Файлов выписок прочитано: {bank_file_count} из {len(bank_files)}")
        print(f"Банковских платежей отобрано: {len(bank_rows)}")
        print(f"Плановых групп: {len(groups)}")
        print(f"Найдено совпадений и перенесено групп: {appended_groups}")
        print(f"Перенесено строк на '{RESULT_SHEET}': {total_plan_rows}")
        print(f"Физически удалено строк с '{PLAN_SHEET}': {len(rows_to_delete)}")
        print(f"Групп без совпадения по сумме и городу: {no_match_count}")
        print(
            "Групп с несколькими подходящими платежами или конфликтом платежа: "
            f"{ambiguous_count}"
        )
        print(f"Групп, уже находившихся на '{RESULT_SHEET}': {already_on_send_count}")
        print(f"Конфликтов суммы между data и send: {send_conflict_count}")
        print(
            "Групп, претендовавших на одну банковскую строку: "
            f"{bank_collision_count}"
        )
        print(f"Файл сохранён: {plan_file}")
        print(f"Резервная копия: {backup}")
        print("=" * 68)

    finally:
        if wb_values is not None:
            try:
                wb_values.close()
            except Exception:
                pass
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if temp_dir is not None and temp_dir.exists():
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nОШИБКА: {exc}")
    input("\nНажмите Enter для выхода...")
