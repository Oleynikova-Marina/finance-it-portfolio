"""Создание черновика письма Outlook с подготовленными выписками во вложении.

Демонстрационный пример: заполните адрес отправителя и список получателей
своими тестовыми/рабочими настройками локально. Не публикуйте реальные адреса.
Письмо только сохраняется в «Черновиках» — автоматической отправки нет.
"""
from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import win32com.client


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
SENDER_EMAIL = "sender@example.com"
RECIPIENTS = ["colleague1@example.com", "colleague2@example.com"]
SUBJECT_PREFIX = "Банковская выписка"


def collect_attachments(folder: Path) -> list[Path]:
    if not folder.is_dir():
        raise FileNotFoundError(f"Папка с подготовленными файлами не найдена: {folder}")

    return sorted(
        path for path in folder.iterdir()
        if path.is_file()
        and path.suffix.casefold() in EXCEL_EXTENSIONS
        and not path.name.startswith("~$")
    )


def find_sender_account(namespace, sender_email: str):
    target = sender_email.strip().casefold()
    for account in namespace.Accounts:
        address = str(getattr(account, "SmtpAddress", "") or "").strip().casefold()
        if address == target:
            return account
    raise RuntimeError(
        "В классическом Outlook не найдена указанная учётная запись отправителя. "
        "Проверьте локальную настройку Outlook."
    )


def create_draft(attachments: list[Path], sender_email: str, recipients: list[str]) -> None:
    if not attachments:
        raise ValueError("Нет Excel-файлов для вложения.")
    if not recipients:
        raise ValueError("Список получателей пуст.")

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    sender_account = find_sender_account(namespace, sender_email)

    mail = outlook.CreateItem(0)  # olMailItem
    mail.SendUsingAccount = sender_account
    mail.To = "; ".join(recipients)
    mail.Subject = f"{SUBJECT_PREFIX} на {date.today():%d.%m.%Y}"
    mail.Body = (
        "Здравствуйте!\r\n"
        "\r\n"
        "\r\n"
        "\r\n"
        "С уважением,\r\n"
        "Сотрудник\r\n"
        "ФИО\r\n"
        "Телефон"
    )

    for path in attachments:
        mail.Attachments.Add(str(path.resolve()))

    # Намеренно отсутствует вызов mail.Send(): письмо остаётся черновиком.
    mail.Save()
    print(
        f"Черновик сохранён. Вложений: {len(attachments)}. "
        "Письмо не отправлено — проверьте его в Outlook."
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Создать черновик Outlook с подготовленными Excel-выписками."
    )
    parser.add_argument("--attachments", type=Path, required=True)
    args = parser.parse_args()

    files = collect_attachments(args.attachments)
    create_draft(files, SENDER_EMAIL, RECIPIENTS)


if __name__ == "__main__":
    main()
