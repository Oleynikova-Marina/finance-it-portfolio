from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
CASES = {
    "BANK1": {
        "source": ROOT / "BANK1" / "source" / "BANK1_statement_14.11.2026.xlsx",
        "output": ROOT / "BANK1" / "output" / "BANK1_cleaned_14.11.2026.xlsx",
    },
    "BANK2": {
        "source": ROOT / "BANK2" / "source" / "BANK2_statement_14.11.2026.xlsx",
        "output": ROOT / "BANK2" / "output" / "BANK2_cleaned_14.11.2026.xlsx",
    },
}


def find_text_cell(ws, text):
    target = text.casefold()
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and target in str(cell.value).casefold():
                return cell
    return None


def header_columns(ws, required, scan_rows=30):
    required_norm = {item.casefold() for item in required}
    for row_num in range(1, min(ws.max_row, scan_rows) + 1):
        values = {
            str(ws.cell(row_num, col_num).value).strip().casefold(): col_num
            for col_num in range(1, ws.max_column + 1)
            if ws.cell(row_num, col_num).value is not None
        }
        if required_norm.issubset(values):
            return row_num, values
    raise AssertionError(f"Не найдена строка заголовков: {required}")


def check_xlsx_integrity(path):
    with ZipFile(path) as archive:
        assert archive.testzip() is None, f"Повреждён XLSX: {path.name}"


def check_bank1():
    paths = CASES["BANK1"]
    assert paths["source"].is_file(), "BANK1: исходная книга не найдена"
    assert paths["output"].is_file(), "BANK1: результат не найден"
    source_wb = load_workbook(paths["source"], data_only=True)
    output_wb = load_workbook(paths["output"], data_only=True)
    checks = {}
    try:
        src, out = source_wb.active, output_wb.active
        checks["source_file_present"] = paths["source"].is_file()
        checks["output_file_present"] = paths["output"].is_file()
        checks["source_and_output_open"] = src.title == out.title == "Выписка"

        opening = find_text_cell(out, "Входящий остаток на начало дня")
        checks["opening_balance_cell_cleared"] = (
            opening is not None and out.cell(opening.row + 1, opening.column).value is None
        )

        header_row, cols = header_columns(
            out, ("Номер документа", "Дебет / Debit", "Кредит / Credit")
        )
        doc_col, debit_col = cols["номер документа"], cols["дебет / debit"]
        operation_rows = [
            row_num for row_num in range(header_row + 2, out.max_row + 1)
            if out.cell(row_num, doc_col).value is not None
        ]
        retained = [str(out.cell(row_num, doc_col).value) for row_num in operation_rows]
        checks["only_expected_transactions_retained"] = retained == ["1001", "1004"]
        checks["debit_rows_removed"] = all(
            out.cell(row_num, debit_col).value in (None, "")
            for row_num in operation_rows
        )
        checks["internal_transfer_removed"] = (
            find_text_cell(out, "Перевод собственных средств") is None
        )
        checks["operations_count_row_removed"] = (
            find_text_cell(out, "Количество операций") is None
        )
        checks["totals_removed"] = (
            find_text_cell(out, "Итого дебет") is None
            and find_text_cell(out, "Итого кредит") is None
        )
        checks["merged_header_preserved"] = (
            set(map(str, src.merged_cells.ranges))
            == set(map(str, out.merged_cells.ranges))
        )
        check_xlsx_integrity(paths["source"])
        check_xlsx_integrity(paths["output"])
        checks["xlsx_container_integrity"] = True
    finally:
        source_wb.close()
        output_wb.close()
    return checks


def check_bank2():
    paths = CASES["BANK2"]
    assert paths["source"].is_file(), "BANK2: исходная книга не найдена"
    assert paths["output"].is_file(), "BANK2: результат не найден"
    source_wb = load_workbook(paths["source"], data_only=True)
    output_wb = load_workbook(paths["output"], data_only=True)
    checks = {}
    try:
        src, out = source_wb.active, output_wb.active
        checks["source_file_present"] = paths["source"].is_file()
        checks["output_file_present"] = paths["output"].is_file()
        checks["source_and_output_open"] = src.title == out.title == "Выписка"

        outgoing_source = find_text_cell(src, "Исходящее сальдо")
        outgoing_value = (
            src.cell(outgoing_source.row, outgoing_source.column + 1).value
            if outgoing_source else None
        )
        checks["source_outgoing_balance_corrected"] = outgoing_value == 1_267_800

        header_row, src_cols = header_columns(src, ("Дебет", "Кредит", "Дата", "Документ"))
        debit_col, credit_col = src_cols["дебет"], src_cols["кредит"]
        opening = find_text_cell(src, "Входящее сальдо")
        opening_value = (
            src.cell(opening.row, opening.column + 1).value if opening else None
        )
        outgoing_row = outgoing_source.row if outgoing_source else src.max_row + 1
        credits = sum(
            float(src.cell(row_num, credit_col).value or 0)
            for row_num in range(header_row + 1, outgoing_row)
        )
        debits = sum(
            float(src.cell(row_num, debit_col).value or 0)
            for row_num in range(header_row + 1, outgoing_row)
        )
        checks["source_balance_reconciles"] = (
            opening_value is not None
            and round(float(opening_value) + credits - debits, 2)
            == round(float(outgoing_value), 2)
        )

        incoming_output = find_text_cell(out, "Входящее сальдо")
        checks["opening_balance_cell_cleared"] = (
            incoming_output is not None
            and out.cell(
                incoming_output.row, incoming_output.column + 1
            ).value is None
        )

        out_header_row, out_cols = header_columns(out, ("Дебет", "Кредит", "Документ"))
        out_debit = out_cols["дебет"]
        out_credit = out_cols["кредит"]
        out_doc = out_cols["документ"]
        operation_rows = [
            row_num for row_num in range(out_header_row + 1, out.max_row + 1)
            if out.cell(row_num, out_doc).value is not None
        ]
        retained = [str(out.cell(row_num, out_doc).value) for row_num in operation_rows]
        checks["only_expected_transactions_retained"] = retained == ["ПЛ-2001", "ПЛ-2004"]
        checks["retained_rows_are_credit_only"] = all(
            out.cell(row_num, out_debit).value in (None, "")
            and out.cell(row_num, out_credit).value not in (None, "")
            for row_num in operation_rows
        )
        checks["internal_transfer_removed"] = (
            find_text_cell(out, "Перевод средств между своими счетами") is None
        )
        checks["outgoing_balance_row_removed"] = (
            find_text_cell(out, "Исходящее сальдо") is None
        )
        checks["rows_after_outgoing_removed"] = (
            find_text_cell(out, "Строка после исходящего сальдо") is None
            and find_text_cell(out, "Дополнительная служебная информация") is None
        )
        check_xlsx_integrity(paths["source"])
        check_xlsx_integrity(paths["output"])
        checks["xlsx_container_integrity"] = True
    finally:
        source_wb.close()
        output_wb.close()
    return checks


def write_report(case, checks):
    report = {
        "case": case,
        "data_type": "synthetic",
        "passed": sum(bool(value) for value in checks.values()),
        "total": len(checks),
        "result": "PASS" if all(checks.values()) else "FAIL",
        "checks": checks,
    }
    folder = ROOT / case
    folder.mkdir(exist_ok=True)
    (folder / "qa_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main():
    results = {"BANK1": check_bank1(), "BANK2": check_bank2()}
    for case, checks in results.items():
        write_report(case, checks)
        count = sum(bool(value) for value in checks.values())
        print(f"{case}: {count}/{len(checks)} — {'PASS' if all(checks.values()) else 'FAIL'}")
        for name, passed in checks.items():
            if not passed:
                print(f"  FAIL: {name}")

    passed = sum(sum(bool(value) for value in checks.values()) for checks in results.values())
    total = sum(len(checks) for checks in results.values())
    print(f"TOTAL: {passed}/{total} — {'PASS' if passed == total else 'FAIL'}")
    summary = {
        "project": "Synthetic bank statement cleanup demo",
        "all_data_synthetic": True,
        "cases": {
            case: {
                "passed": sum(bool(value) for value in checks.values()),
                "total": len(checks),
                "result": "PASS" if all(checks.values()) else "FAIL",
            }
            for case, checks in results.items()
        },
        "overall_passed": passed,
        "overall_total": total,
        "overall_result": "PASS" if passed == total else "FAIL",
    }
    (ROOT / "qa_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    if passed != total:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
