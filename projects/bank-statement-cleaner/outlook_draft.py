"""Создание черновика Outlook с обработанными выписками.
Настройки адресов и подписи находятся в локальном config.py.
Письмо только сохраняется в «Черновиках» — автоматической отправки нет.
Требуются классический Outlook для Windows и pywin32.
"""
from datetime import date
from pathlib import Path
import win32com.client
from config import OUTPUT_DIR, OUTLOOK_SENDER, OUTLOOK_RECIPIENTS, SIGNATURE_NAME, SIGNATURE_PHONE


def find_sender_account(namespace, address):
    target = address.strip().casefold()
    for account in namespace.Accounts:
        actual = str(getattr(account, "SmtpAddress", "") or "").strip().casefold()
        if actual == target:
            return account
    raise RuntimeError("Указанная учётная запись не найдена в Outlook.")


def create_draft():
    folder = Path(OUTPUT_DIR)
    if not folder.is_dir():
        raise FileNotFoundError(f"Папка результатов не найдена: {folder}")
    files = sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.casefold() in {".xlsx", ".xlsm", ".xls"}
        and not p.name.startswith("~$")
    )
    if not files:
        raise RuntimeError("Нет Excel-файлов для вложения.")
    if not OUTLOOK_RECIPIENTS:
        raise RuntimeError("В локальном config.py не заданы получатели.")

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    mail = outlook.CreateItem(0)
    mail.SendUsingAccount = find_sender_account(namespace, OUTLOOK_SENDER)
    mail.To = "; ".join(OUTLOOK_RECIPIENTS)
    mail.Subject = f"Выписка на {date.today():%d.%m.%Y}"
    mail.HTMLBody = f"""<html><body><div style="font-family:Georgia,serif;font-size:12pt;color:#000000">
    Здравствуйте!<br><br><br><br>
    С уважением,<br>{SIGNATURE_NAME}<br>{SIGNATURE_PHONE}
    </div></body></html>"""
    for path in files:
        mail.Attachments.Add(str(path.resolve()))
    mail.Save()
    print(f"Черновик сохранён. Вложений: {len(files)}. Письмо НЕ отправлено.")


if __name__ == "__main__":
    create_draft()
