"""Единый последовательный запуск двух обработчиков и Outlook draft."""
from config import BANK_1_INPUT, BANK_2_INPUT, OUTPUT_DIR
from bank_1_cleaner import process_folder as process_bank_1
from bank_2_cleaner import process_folder as process_bank_2


def main():
    print("\n=== БАНК 1 ===")
    n1 = process_bank_1(BANK_1_INPUT, OUTPUT_DIR)
    print("\n=== БАНК 2 ===")
    n2 = process_bank_2(BANK_2_INPUT, OUTPUT_DIR)
    print("\n=== РЕЗУЛЬТАТ ===")
    print(f"Обработано файлов: Банк 1 — {n1}; Банк 2 — {n2}")

    if n1 + n2:
        answer = input("Создать черновик Outlook с файлами из общей папки результатов? [д/н]: ").strip().casefold()
        if answer in {"д", "да", "y", "yes"}:
            from outlook_draft import create_draft
            create_draft()
        else:
            print("Черновик не создавался.")
    else:
        print("Нет обработанных файлов. Черновик не создавался.")


if __name__ == "__main__":
    main()
