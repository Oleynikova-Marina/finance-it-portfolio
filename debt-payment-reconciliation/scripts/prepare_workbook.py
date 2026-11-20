from pathlib import Path
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
SOURCE_FOLDER = BASE_DIR / "input"
TARGET_FOLDER = BASE_DIR / "working"

DATA_START_ROW = 3
NUMBER_COLUMNS = (10, 12)
HIDDEN_COLUMNS = (3, 11, 13)
GROUP_COLUMN = 2
RS_CITY_COLUMN = 14  # N — поле «рс», из которого берётся название города

# Более тёмный терракотовый оттенок
DARK_TERRACOTTA = "8B3F2F"

# Дополнительный блок дашборда. Для отключения установите False.
ENABLE_TREEMAP_DASHBOARD = False
TREEMAP_RС_COLUMN = 14       # N — столбец «рс»
TREEMAP_AMOUNT_COLUMN = 12   # L — столбец с суммой
TREEMAP_DATA_START_ROW = DATA_START_ROW
TREEMAP_LEVELS = 5
TREEMAP_GRID_COLUMNS = 62
TREEMAP_COL_WIDTH = 1.87
TREEMAP_COLORS = [
    "443A43",  # фиолетовый
    "C08081",  # викторианский розовый
    "3949AB",  # ультрамарин
    "C17D11",  # охра
    "2E7D5B",  # зелёный
    "168C8C",  # бирюзовый
    "B23A48",  # красный
    "7B6D8D",  # приглушённый лиловый
    "52796F",  # серо-зелёный
    "5C677D",  # сине-серый
    "A64D79",  # насыщенный розово-сливовый
    "D17A52",  # терракотовый
]


def to_number(value):
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value

    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        number = float(text)
        return int(number) if number.is_integer() else number
    except ValueError:
        return value


def sort_key(value):
    if value is None or str(value).strip() == "":
        return (1, "")
    return (0, str(value).strip().casefold())



def clean_rs_city(value):
    """Нормализует город из поля рс: удаляет служебные маркеры ОБЩИЙ и СОСЕНКИ."""
    if value is None:
        return ""
    import re

    text = " ".join(str(value).replace("\u00a0", " ").split()).strip()
    text = re.sub(r"^\s*ОБЩИЙ\s*[,;:—–-]?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"(?:\s*[-–—]\s*|\s+)СОСЕНКИ\s*$", "", text, flags=re.IGNORECASE)
    text = text.strip(" ,;:-—–")
    return text.upper()


def make_treemap_dashboard(workbook, data_ws):
    """Создаёт лист treemap в виде пяти цветных горизонтальных ярусов."""
    from math import floor

    grouped = {}
    for row_idx in range(TREEMAP_DATA_START_ROW, data_ws.max_row + 1):
        raw_city = data_ws.cell(row=row_idx, column=TREEMAP_RС_COLUMN).value
        city = clean_rs_city(raw_city)
        amount = to_number(data_ws.cell(row=row_idx, column=TREEMAP_AMOUNT_COLUMN).value)
        if not city or not isinstance(amount, (int, float)) or isinstance(amount, bool):
            continue
        grouped[city] = grouped.get(city, 0.0) + float(amount)

    grouped = {city: total for city, total in grouped.items() if total > 0}
    if "treemap" in workbook.sheetnames:
        del workbook["treemap"]
    # Удаляем старый лист, чтобы после обновления не оставался дубль.
    if "dashboard" in workbook.sheetnames:
        del workbook["dashboard"]
    dash = workbook.create_sheet("treemap")
    dash.sheet_view.showGridLines = False
    dash.freeze_panes = "A3"

    for col in range(1, TREEMAP_GRID_COLUMNS + 1):
        dash.column_dimensions[dash.cell(row=1, column=col).column_letter].width = TREEMAP_COL_WIDTH

    dash.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TREEMAP_GRID_COLUMNS)
    title = dash.cell(1, 1, "ДАННЫЕ ПО ЗАДОЛЖЕННОСТИ ПО ГОРОДАМ")
    title.font = Font(name="Times New Roman", size=16, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="29243A")
    title.alignment = Alignment(horizontal="center", vertical="center")
    dash.row_dimensions[1].height = 30

    if not grouped:
        dash.merge_cells(start_row=3, start_column=1, end_row=4, end_column=TREEMAP_GRID_COLUMNS)
        cell = dash.cell(3, 1, "Нет данных для построения диаграммы: проверьте столбцы N (рс) и L (сумма).")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.font = Font(name="Times New Roman", size=12, color="666666")
        return

    items = sorted(grouped.items(), key=lambda item: item[1], reverse=True)

    # Москва и Санкт-Петербург — отдельный верхний ярус.
    # Остальные города равномерно распределяются по четырём ярусам ниже.
    top_cities = items[:2]
    remaining = items[2:]
    tiers = [top_cities]
    lower_levels = 4
    for level in range(lower_levels):
        left = len(remaining) * level // lower_levels
        right = len(remaining) * (level + 1) // lower_levels
        if right > left:
            tiers.append(remaining[left:right])
    tiers = [tier for tier in tiers if tier]

    # Фиксированная высота ярусов вместо прежней диаграммы на ~97 строк:
    # Первые два яруса — по 6 строк; остальные — 4, 3 и 3 строки.
    requested_heights = [6, 6, 4, 3, 3]
    tier_heights = requested_heights[:len(tiers)]

    current_row = 3
    color_idx = 0
    for tier_idx, (tier, tier_height) in enumerate(zip(tiers, tier_heights)):
        tier_sum = sum(v for _, v in tier)
        # Целочисленное распределение колонок с гарантией хотя бы одной колонки на плитку.
        widths = [max(1, round(TREEMAP_GRID_COLUMNS * v / tier_sum)) for _, v in tier]
        while sum(widths) > TREEMAP_GRID_COLUMNS:
            candidates = [i for i, w in enumerate(widths) if w > 1]
            if not candidates:
                break
            j = max(candidates, key=lambda i: widths[i])
            widths[j] -= 1
        while sum(widths) < TREEMAP_GRID_COLUMNS:
            j = max(range(len(tier)), key=lambda i: tier[i][1] / widths[i])
            widths[j] += 1

        end_row = current_row + tier_height - 1
        for (city, amount), width in zip(tier, widths):
            start_col = 1 if not any(dash.cell(r, 1).value for r in range(current_row, end_row + 1)) else 1
            # Плитки размещаются последовательно слева направо в пределах яруса.
            if 'next_col' not in locals():
                next_col = 1
            start_col = next_col
            end_col = min(TREEMAP_GRID_COLUMNS, start_col + width - 1)
            if start_col <= end_col:
                dash.merge_cells(start_row=current_row, start_column=start_col, end_row=end_row, end_column=end_col)
                cell = dash.cell(current_row, start_col)
                thousands = amount / 1000
                amount_label = f"{thousands:,.0f}".replace(",", " ")
                cell.value = f"{city.upper()}\n{amount_label}т.р."
                color = TREEMAP_COLORS[color_idx % len(TREEMAP_COLORS)]
                color_idx += 1
                cell.fill = PatternFill("solid", fgColor=color)
                is_top_city = tier_idx == 0 and city in {name for name, _ in top_cities}
                # Начальный размер увеличен на 30%; для каждой плитки подбираем
                # максимальный размер, который ориентировочно помещает обе строки.
                base_font_size = 29 if is_top_city else (9 if width >= 6 else 8)
                tile_width = sum(
                    dash.column_dimensions[get_column_letter(c)].width or TREEMAP_COL_WIDTH
                    for c in range(start_col, end_col + 1)
                )
                tile_height = sum(
                    dash.row_dimensions[r].height or 16
                    for r in range(current_row, end_row + 1)
                )
                city_text = city.upper()
                amount_text = f"{amount_label}т.р."
                font_size = base_font_size
                while font_size > 5:
                    # Приближённая оценка доступной ширины и высоты Excel.
                    chars_per_line = max(3, int(tile_width * 1.05 * 72 / (font_size * 0.58)))
                    lines_needed = max(
                        (len(city_text) + chars_per_line - 1) // chars_per_line,
                        (len(amount_text) + chars_per_line - 1) // chars_per_line,
                    ) + 1
                    if lines_needed * font_size * 1.25 <= tile_height * 0.9:
                        break
                    font_size -= 1
                cell.font = Font(name="Times New Roman", size=font_size, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(
                    left=Side(style="medium", color="FFFFFF"),
                    right=Side(style="medium", color="FFFFFF"),
                    top=Side(style="medium", color="FFFFFF"),
                    bottom=Side(style="medium", color="FFFFFF"),
                )
            next_col = end_col + 1
        current_row = end_row + 1
        if 'next_col' in locals():
            del next_col

    for r in range(3, current_row):
        dash.row_dimensions[r].height = 21
    dash.sheet_properties.pageSetUpPr.fitToPage = True
    dash.page_setup.fitToWidth = 1
    dash.page_setup.fitToHeight = 1
    dash.sheet_properties.outlinePr.summaryBelow = True
    print(f"Лист treemap создан: {len(items)} городов, {len(tiers)} ярусов.")


def transfer_source_data(source_path, target_ws):
    """Переносит шапку из строки 2 и данные с 3-й строки; очищает старый хвост."""
    from copy import copy

    source_wb = load_workbook(
        source_path,
        data_only=False,
        keep_vba=source_path.suffix.casefold() == ".xlsm",
    )
    try:
        source_ws = source_wb.active
        source_last_row = source_ws.max_row
        source_last_col = source_ws.max_column
        target_last_row = target_ws.max_row
        target_last_col = target_ws.max_column
        last_col = max(source_last_col, target_last_col)

        # Переносим всю шапку из строки 2: значения, оформление и размеры столбцов.
        for col_idx in range(1, last_col + 1):
            src = source_ws.cell(row=2, column=col_idx)
            dst = target_ws.cell(row=2, column=col_idx)
            if not isinstance(dst, MergedCell):
                dst.value = src.value if col_idx <= source_last_col else None
                if col_idx <= source_last_col:
                    if src.has_style:
                        dst._style = copy(src._style)
                    if src.number_format:
                        dst.number_format = src.number_format
                    if src.alignment:
                        dst.alignment = copy(src.alignment)
                    if src.protection:
                        dst.protection = copy(src.protection)
                    if src.comment:
                        dst.comment = copy(src.comment)
                    else:
                        dst.comment = None
            letter = target_ws.cell(row=2, column=col_idx).column_letter
            if col_idx <= source_last_col:
                source_letter = source_ws.cell(row=2, column=col_idx).column_letter
                target_ws.column_dimensions[letter].width = source_ws.column_dimensions[source_letter].width
                target_ws.column_dimensions[letter].hidden = source_ws.column_dimensions[source_letter].hidden

        target_ws.row_dimensions[2].height = source_ws.row_dimensions[2].height

        # Очищаем старые данные и оформление с 3-й строки до конца прежнего диапазона.
        old_end = max(source_last_row, target_last_row)
        for row_idx in range(DATA_START_ROW, old_end + 1):
            for col_idx in range(1, last_col + 1):
                cell = target_ws.cell(row=row_idx, column=col_idx)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = None
                # Убираем заливку/границы/шрифт старых строк, чтобы не оставался серый хвост.
                if row_idx > source_last_row:
                    cell._style = None
                    cell.number_format = "General"
                    cell.comment = None

        # Переносим данные начиная с 3-й строки.
        copied_rows = 0
        for row_idx in range(DATA_START_ROW, source_last_row + 1):
            row_has_data = False
            for col_idx in range(1, source_last_col + 1):
                value = source_ws.cell(row=row_idx, column=col_idx).value
                target_cell = target_ws.cell(row=row_idx, column=col_idx)
                if not isinstance(target_cell, MergedCell):
                    target_cell.value = value
                # Копируем оформление исходной ячейки вместе со значением,
                # особенно границы, чтобы сетка Excel не пропадала после переноса.
                if not isinstance(target_cell, MergedCell):
                    source_cell = source_ws.cell(row=row_idx, column=col_idx)
                    if source_cell.has_style:
                        target_cell._style = copy(source_cell._style)
                    if source_cell.number_format:
                        target_cell.number_format = source_cell.number_format
                    if source_cell.alignment:
                        target_cell.alignment = copy(source_cell.alignment)
                    if source_cell.protection:
                        target_cell.protection = copy(source_cell.protection)
                    if source_cell.comment:
                        target_cell.comment = copy(source_cell.comment)
                    else:
                        target_cell.comment = None
                if value is not None and str(value).strip():
                    row_has_data = True
            if row_has_data:
                copied_rows += 1

        # Убираем высоты строк, оставшиеся от старого длинного отчёта.
        for row_idx in range(source_last_row + 1, target_last_row + 1):
            target_ws.row_dimensions[row_idx].height = None
            target_ws.row_dimensions[row_idx].hidden = False

        return copied_rows
    finally:
        source_wb.close()


def process_file(source_path, target_path):
    workbook = load_workbook(
        target_path,
        keep_vba=target_path.suffix.casefold() == ".xlsm"
    )

    if "data" not in workbook.sheetnames:
        raise ValueError(
            "В целевом файле не найден лист 'data'. "
            "Добавьте лист с таким названием и запустите скрипт снова."
        )
    ws = workbook["data"]

    # 0. Перенести свежие данные из исходного файла, сохранив первые две строки.
    copied_rows = transfer_source_data(source_path, ws)
    print(f"Перенесено строк из исходного файла: {copied_rows}", flush=True)

    # 1. Преобразовать значения в столбцах 10 и 12 в числа.
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        for col_idx in NUMBER_COLUMNS:
            cell = ws.cell(row=row_idx, column=col_idx)
            if not isinstance(cell, MergedCell):
                cell.value = to_number(cell.value)

    # 2. Очистить названия городов в столбце «рс» от служебных маркеров.
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=RS_CITY_COLUMN)
        if not isinstance(cell, MergedCell) and cell.value is not None:
            cell.value = clean_rs_city(cell.value)

    # 3. Скрыть столбцы 3, 11 и 13.
    for col_idx in HIDDEN_COLUMNS:
        letter = ws.cell(row=2, column=col_idx).column_letter
        ws.column_dimensions[letter].hidden = True

    # 4. Сгруппировать строки по значениям столбца 2 в алфавитном порядке.
    last_row = ws.max_row
    last_col = ws.max_column
    if last_row >= DATA_START_ROW:
        rows = [
            [ws.cell(row=r, column=c).value for c in range(1, last_col + 1)]
            for r in range(DATA_START_ROW, last_row + 1)
        ]
        rows.sort(key=lambda row: sort_key(row[GROUP_COLUMN - 1]))

        for r_idx, values in enumerate(rows, start=DATA_START_ROW):
            for c_idx, value in enumerate(values, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if not isinstance(cell, MergedCell):
                    cell.value = value

    # 5. Сумма столбца 12 в L1.
    total = sum(
        ws.cell(row=r, column=12).value
        for r in range(DATA_START_ROW, ws.max_row + 1)
        if isinstance(ws.cell(row=r, column=12).value, (int, float))
        and not isinstance(ws.cell(row=r, column=12).value, bool)
    )
    result_cell = ws["L1"]
    result_cell.value = total

    # Формат Excel с группировкой разрядов.
    # Не заменяйте запятую-разделитель групп на пробел в коде формата:
    # Excel сам отображает разделитель согласно системным настройкам.
    # Для русского регионального формата это обычно пробелы и десятичная запятая.
    result_cell.number_format = '#,##0.00;[Red]-#,##0.00'
    result_cell.font = Font(
        name="Times New Roman",
        size=14,
        bold=True,
        color=DARK_TERRACOTTA
    )

    # Автоматически расширяем столбец L, чтобы сумма в L1 отображалась целиком.
    # Оставляем небольшой запас, не уменьшая уже заданную ширину столбца.
    # Ширину считаем по фактическому отображению суммы в Excel:
    # разряды, десятичная часть, знак минуса и запас на шрифт.
    display_total = f"{total:,.2f}"
    required_width = min(255, max(14, len(display_total) + 5))
    ws.column_dimensions["L"].width = required_width

    # ДОПОЛНИТЕЛЬНЫЙ БЛОК: treemap можно отключить через ENABLE_TREEMAP_DASHBOARD = False.
    if ENABLE_TREEMAP_DASHBOARD:
        make_treemap_dashboard(workbook, ws)

    # Повторно фиксируем скрытие ПОСЛЕ всех операций с листом и перед сохранением,
    # чтобы копирование размеров столбцов из исходного файла не отменяло настройку.
    for col_idx in HIDDEN_COLUMNS:
        letter = get_column_letter(col_idx)
        ws.column_dimensions[letter].hidden = True

    # Если в книге есть объединение, затрагивающее L1, Excel не сможет нормально
    # показать сумму в обычной ячейке — сообщаем об этом явно в консоль.
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= 1 <= merged_range.max_row and merged_range.min_col <= 12 <= merged_range.max_col:
            print(f"Предупреждение: L1 входит в объединённый диапазон {merged_range}; ширина столбца L может не влиять на отображение.", flush=True)
            break

    workbook.save(target_path)
    print(f"Готово. Файл обновлён: {target_path}")
    print(f"Сумма столбца 12: {total:,.2f}".replace(",", " "))


def main():
    print("=== ПОДГОТОВКА РАБОЧЕЙ КНИГИ ===", flush=True)
    print(f"Источник: {SOURCE_FOLDER}", flush=True)
    print(f"Назначение: {TARGET_FOLDER}", flush=True)

    if not SOURCE_FOLDER.exists():
        print(f"Папка с исходным файлом не найдена: {SOURCE_FOLDER}")
        return
    if not TARGET_FOLDER.exists():
        print(f"Папка с целевым файлом не найдена: {TARGET_FOLDER}")
        return

    source_files = [
        p for p in SOURCE_FOLDER.iterdir()
        if p.is_file()
        and p.suffix.casefold() in {".xlsx", ".xlsm"}
        and not p.name.startswith("~$")
    ]
    target_files = [
        p for p in TARGET_FOLDER.iterdir()
        if p.is_file()
        and p.suffix.casefold() in {".xlsx", ".xlsm"}
        and not p.name.startswith("~$")
    ]

    if len(source_files) != 1:
        print(f"В папке input ожидался один Excel-файл, найдено: {len(source_files)}.")
        return
    if len(target_files) != 1:
        print(f"В папке working ожидался один Excel-файл, найдено: {len(target_files)}.")
        return

    print(f"Исходный файл: {source_files[0].name}", flush=True)
    print(f"Целевой файл: {target_files[0].name}", flush=True)

    try:
        process_file(source_files[0], target_files[0])
    except PermissionError:
        print("Не удалось сохранить файл. Закройте его в Excel и запустите скрипт снова.", flush=True)
    except Exception as exc:
        import traceback
        print(f"Ошибка обработки: {exc}", flush=True)
        traceback.print_exc()


if __name__ == "__main__":
    main()