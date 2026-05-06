import sys

VAT_RATE = 0.22      # 22%
VAT_SHARE = 22 / 122 # для выделения НДС из суммы с НДС

def format_rub(value: float) -> str:
    return f"{value:,.0f} ₽".replace(",", " ")

def calc_add_vat():
    try:
        base = float(input("Введите сумму без НДС (руб): ").replace(",", "."))
    except ValueError:
        print("Некорректное число.")
        return

    vat = base * VAT_RATE
    total = base + vat

    print("\n--- Результат ---")
    print(f"Сумма без НДС: {format_rub(base)}")
    print(f"НДС 22%:       {format_rub(vat)}")
    print(f"Итого с НДС:   {format_rub(total)}")
    print("-----------------\n")

def calc_extract_vat():
    try:
        total = float(input("Введите сумму с НДС (руб): ").replace(",", "."))
    except ValueError:
        print("Некорректное число.")
        return

    vat = total * VAT_SHARE
    base = total - vat

    print("\n--- Результат ---")
    print(f"Сумма с НДС:   {format_rub(total)}")
    print(f"НДС (22/122):  {format_rub(vat)}")
    print(f"Сумма без НДС: {format_rub(base)}")
    print("-----------------\n")

def calc_advance():
    try:
        advance_total = float(input("Аванс с НДС (руб): ").replace(",", "."))
        rest_base = float(input("Остаток без НДС (руб): ").replace(",", "."))
    except ValueError:
        print("Некорректное число.")
        return

    adv_vat = advance_total * VAT_SHARE
    adv_base = advance_total - adv_vat

    rest_vat = rest_base * VAT_RATE
    rest_total = rest_base + rest_vat

    total_vat = adv_vat + rest_vat

    print("\n--- Аванс ---")
    print(f"Аванс с НДС:   {format_rub(advance_total)}")
    print(f"НДС с аванса:  {format_rub(adv_vat)}")
    print(f"База аванса:   {format_rub(adv_base)}")

    print("\n--- Отгрузка ---")
    print(f"Остаток без НДС: {format_rub(rest_base)}")
    print(f"НДС 22%:         {format_rub(rest_vat)}")
    print(f"Счет на отгрузку:{format_rub(rest_total)}")

    print("\n--- Итог по договору ---")
    print(f"Всего НДС:       {format_rub(total_vat)}")
    print("-------------------------\n")

def main():
    while True:
        print("Калькулятор НДС 2026 (ставка 22%)")
        print("1 – Начислить НДС (от суммы без НДС)")
        print("2 – Выделить НДС (из суммы с НДС)")
        print("3 – Авансовые платежи")
        print("0 – Выход")

        choice = input("Выберите режим: ").strip()

        if choice == "1":
            calc_add_vat()
        elif choice == "2":
            calc_extract_vat()
        elif choice == "3":
            calc_advance()
        elif choice == "0":
            print("Выход.")
            break
        else:
            print("Некорректный выбор.\n")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nЗавершение работы.")
        sys.exit(0)
