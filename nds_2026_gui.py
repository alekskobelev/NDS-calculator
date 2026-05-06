import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

VAT_RATE = 0.22       # 22%
VAT_SHARE = 22 / 122  # для выделения НДС из суммы с НДС


def format_rub(value: float) -> str:
    """Форматирование числа в рубли с пробелами."""
    try:
        return f"{value:,.0f} ₽".replace(",", " ")
    except Exception:
        return "—"


def copy_to_clipboard(window: tk.Tk, text: str):
    """Копирование текста в буфер обмена."""
    # убираем символ ₽ и пробелы, если нужно копировать только число
    clean = text.replace("₽", "").strip()
    window.clipboard_clear()
    window.clipboard_append(clean)
    window.update()


class NDSApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Калькулятор НДС 2026 (22%)")
        self.geometry("560x460")
        self.resizable(False, False)

        # Основной контейнер
        container = ttk.Frame(self, padding=10)
        container.pack(fill="both", expand=True)

        # Notebook с вкладками
        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)

        self.tab_add = ttk.Frame(notebook, padding=10)
        self.tab_extract = ttk.Frame(notebook, padding=10)
        self.tab_advance = ttk.Frame(notebook, padding=10)

        notebook.add(self.tab_add, text="НДС начислить")
        notebook.add(self.tab_extract, text="НДС выделить")
        notebook.add(self.tab_advance, text="Авансы")

        self.build_tab_add()
        self.build_tab_extract()
        self.build_tab_advance()

    # ---------- вкладка 1: начислить НДС ----------
    def build_tab_add(self):
        ttk.Label(
            self.tab_add,
            text="Расчет НДС при добавлении 22% к сумме без НДС",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", pady=(0, 5))

        ttk.Label(self.tab_add, text="Введите сумму без НДС (руб):").pack(anchor="w")

        self.entry_add_base = ttk.Entry(self.tab_add)
        self.entry_add_base.pack(fill="x", pady=5)
        self.entry_add_base.insert(0, "100000")

        ttk.Button(
            self.tab_add,
            text="Рассчитать",
            command=self.calc_add_vat
        ).pack(pady=5)

        frame_res = ttk.LabelFrame(self.tab_add, text="Результат")
        frame_res.pack(fill="both", expand=True, pady=10)

        row = 0
        ttk.Label(frame_res, text="Сумма без НДС:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_add_base = ttk.Label(frame_res, text="—")
        self.lbl_add_base.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_res,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_add_base.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_res, text="НДС 22%:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_add_vat = ttk.Label(frame_res, text="—")
        self.lbl_add_vat.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_res,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_add_vat.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_res, text="Итого с НДС:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_add_total = ttk.Label(frame_res, text="—", font=("Segoe UI", 10, "bold"))
        self.lbl_add_total.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_res,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_add_total.cget("text"))
        ).grid(row=row, column=2, padx=5)

        frame_res.columnconfigure(1, weight=1)

        ttk.Label(
            self.tab_add,
            text="Формула: НДС = сумма × 0.22; Итого = сумма × 1.22",
            foreground="#555"
        ).pack(anchor="w", pady=(0, 5))

    def calc_add_vat(self):
        try:
            base_str = self.entry_add_base.get().replace(" ", "").replace(",", ".")
            base = float(base_str)
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную сумму без НДС.")
            return

        vat = base * VAT_RATE
        total = base + vat

        self.lbl_add_base.config(text=format_rub(base))
        self.lbl_add_vat.config(text=format_rub(vat))
        self.lbl_add_total.config(text=format_rub(total))

    # ---------- вкладка 2: выделить НДС ----------
    def build_tab_extract(self):
        ttk.Label(
            self.tab_extract,
            text="Выделение НДС 22/122 из суммы с НДС",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", pady=(0, 5))

        ttk.Label(self.tab_extract, text="Введите сумму с НДС (руб):").pack(anchor="w")

        self.entry_ext_total = ttk.Entry(self.tab_extract)
        self.entry_ext_total.pack(fill="x", pady=5)
        self.entry_ext_total.insert(0, "122000")

        ttk.Button(
            self.tab_extract,
            text="Рассчитать",
            command=self.calc_extract_vat
        ).pack(pady=5)

        frame_res = ttk.LabelFrame(self.tab_extract, text="Результат")
        frame_res.pack(fill="both", expand=True, pady=10)

        row = 0
        ttk.Label(frame_res, text="Сумма с НДС:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_ext_total = ttk.Label(frame_res, text="—")
        self.lbl_ext_total.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_res,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_ext_total.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_res, text="НДС (22/122):").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_ext_vat = ttk.Label(frame_res, text="—")
        self.lbl_ext_vat.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_res,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_ext_vat.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_res, text="Сумма без НДС:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_ext_base = ttk.Label(frame_res, text="—", font=("Segoe UI", 10, "bold"))
        self.lbl_ext_base.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_res,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_ext_base.cget("text"))
        ).grid(row=row, column=2, padx=5)

        frame_res.columnconfigure(1, weight=1)

        ttk.Label(
            self.tab_extract,
            text="Формула: НДС = сумма × (22/122); база = сумма − НДС",
            foreground="#555"
        ).pack(anchor="w", pady=(0, 5))

    def calc_extract_vat(self):
        try:
            total_str = self.entry_ext_total.get().replace(" ", "").replace(",", ".")
            total = float(total_str)
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную сумму с НДС.")
            return

        vat = total * VAT_SHARE
        base = total - vat

        self.lbl_ext_total.config(text=format_rub(total))
        self.lbl_ext_vat.config(text=format_rub(vat))
        self.lbl_ext_base.config(text=format_rub(base))

    # ---------- вкладка 3: авансы ----------
    def build_tab_advance(self):
        ttk.Label(
            self.tab_advance,
            text="Расчет НДС по авансам и отгрузке (ставка 22%)",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 5))

        ttk.Label(self.tab_advance, text="Аванс с НДС (руб):").grid(row=1, column=0, sticky="w")
        self.entry_adv_total = ttk.Entry(self.tab_advance)
        self.entry_adv_total.grid(row=2, column=0, sticky="we", padx=(0, 10), pady=5)
        self.entry_adv_total.insert(0, "122000")

        ttk.Label(self.tab_advance, text="Остаток без НДС (руб):").grid(row=1, column=1, sticky="w")
        self.entry_adv_rest = ttk.Entry(self.tab_advance)
        self.entry_adv_rest.grid(row=2, column=1, sticky="we", pady=5)
        self.entry_adv_rest.insert(0, "100000")

        ttk.Button(
            self.tab_advance,
            text="Рассчитать",
            command=self.calc_advance
        ).grid(row=3, column=0, columnspan=2, pady=5)

        # --- блок аванса ---
        frame_adv = ttk.LabelFrame(self.tab_advance, text="Аванс")
        frame_adv.grid(row=4, column=0, sticky="nsew", padx=(0, 10), pady=10)

        row = 0
        ttk.Label(frame_adv, text="Аванс с НДС:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_adv_total = ttk.Label(frame_adv, text="—")
        self.lbl_adv_total.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_adv,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_adv_total.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_adv, text="НДС с аванса:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_adv_vat = ttk.Label(frame_adv, text="—")
        self.lbl_adv_vat.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_adv,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_adv_vat.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_adv, text="База аванса:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_adv_base = ttk.Label(frame_adv, text="—")
        self.lbl_adv_base.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_adv,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_adv_base.cget("text"))
        ).grid(row=row, column=2, padx=5)

        frame_adv.columnconfigure(1, weight=1)

        # --- блок отгрузки ---
        frame_rest = ttk.LabelFrame(self.tab_advance, text="Отгрузка остатка")
        frame_rest.grid(row=4, column=1, sticky="nsew", pady=10)

        row = 0
        ttk.Label(frame_rest, text="Остаток без НДС:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_rest_base = ttk.Label(frame_rest, text="—")
        self.lbl_rest_base.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_rest,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_rest_base.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_rest, text="НДС 22%:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_rest_vat = ttk.Label(frame_rest, text="—")
        self.lbl_rest_vat.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_rest,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_rest_vat.cget("text"))
        ).grid(row=row, column=2, padx=5)

        row += 1
        ttk.Label(frame_rest, text="Счет на отгрузку:").grid(row=row, column=0, sticky="w", pady=2)
        self.lbl_rest_total = ttk.Label(frame_rest, text="—")
        self.lbl_rest_total.grid(row=row, column=1, sticky="w", pady=2)
        ttk.Button(
            frame_rest,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, self.lbl_rest_total.cget("text"))
        ).grid(row=row, column=2, padx=5)

        frame_rest.columnconfigure(1, weight=1)

        # --- итог по договору ---
        frame_total = ttk.LabelFrame(self.tab_advance, text="Итого по договору")
        frame_total.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(0, 5))

        self.lbl_total_vat = ttk.Label(
            frame_total,
            text="Всего НДС (аванс + отгрузка): —",
            font=("Segoe UI", 10, "bold")
        )
        self.lbl_total_vat.grid(row=0, column=0, sticky="w", pady=4)

        ttk.Button(
            frame_total,
            text="Копировать сумму НДС",
            command=lambda: copy_to_clipboard(self, self.lbl_total_vat.cget("text"))
        ).grid(row=0, column=1, padx=5, pady=4, sticky="e")

        self.tab_advance.columnconfigure(0, weight=1)
        self.tab_advance.columnconfigure(1, weight=1)

        ttk.Label(
            self.tab_advance,
            text="Формулы: НДС с аванса = аванс × (22/122); НДС с остатка = остаток × 0.22",
            foreground="#555"
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 5))

    def calc_advance(self):
        try:
            adv_str = self.entry_adv_total.get().replace(" ", "").replace(",", ".")
            rest_str = self.entry_adv_rest.get().replace(" ", "").replace(",", ".")
            advance_total = float(adv_str)
            rest_base = float(rest_str)
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные суммы аванса и остатка.")
            return

        adv_vat = advance_total * VAT_SHARE
        adv_base = advance_total - adv_vat

        rest_vat = rest_base * VAT_RATE
        rest_total = rest_base + rest_vat

        total_vat = adv_vat + rest_vat

        self.lbl_adv_total.config(text=format_rub(advance_total))
        self.lbl_adv_vat.config(text=format_rub(adv_vat))
        self.lbl_adv_base.config(text=format_rub(adv_base))

        self.lbl_rest_base.config(text=format_rub(rest_base))
        self.lbl_rest_vat.config(text=format_rub(rest_vat))
        self.lbl_rest_total.config(text=format_rub(rest_total))

        self.lbl_total_vat.config(
            text=f"Всего НДС (аванс + отгрузка): {format_rub(total_vat)}"
        )


if __name__ == "__main__":
    app = NDSApp()
    app.mainloop()
