# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- ДОБАВЛЕНО: Импорты для экспорта в Excel ---
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# --- Улучшения ---
# 1. Выбор ставки НДС (включая кастомную)
# 2. Автоматический пересчет при вводе и смене ставки
# 3. Кнопка "Очистить" на каждой вкладке
# 4. Рефакторинг кода: убраны повторы, код стал надежнее
# 5. Результаты расчетов теперь нельзя редактировать
# 6. Окно можно растягивать, интерфейс адаптируется
# 7. ДОБАВЛЕНО: Экспорт результатов в Excel (.xlsx)
# -----------------


def format_rub(value: float) -> str:
    """Форматирование числа в рубли с пробелами."""
    try:
        return f"{value:,.2f} ₽".replace(",", " ")
    except (ValueError, TypeError):
        return "—"

def format_rate(rate: float) -> str:
    """Форматирует ставку для отображения, убирая лишние нули."""
    formatted = f"{rate * 100:.2f}".rstrip('0').rstrip('.')
    return formatted

def copy_to_clipboard(window: tk.Tk, text: str):
    """Копирование текста в буфер обмена."""
    clean_text = text.replace("₽", "").replace(" ", "").strip()
    if clean_text == '—':
        return
    window.clipboard_clear()
    window.clipboard_append(clean_text)
    window.update()
    show_tooltip(window, "Скопировано!")


def show_tooltip(widget, text):
    """Показывает временную подсказку рядом с виджетом."""
    x = widget.winfo_rootx() + 25
    y = widget.winfo_rooty() - 25
    
    tooltip = tk.Toplevel(widget)
    tooltip.wm_overrideredirect(True)
    tooltip.wm_geometry(f"+{x}+{y}")
    
    label = tk.Label(tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1,
                   font=("Segoe UI", 8))
    label.pack()
    
    tooltip.after(1000, tooltip.destroy)


class NDSProApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Калькулятор НДС Pro")
        self.geometry("600x580") 
        self.resizable(True, True)
        self.minsize(580, 560)

        self.vat_rates_map = {
            "22% (прогноз)": 0.22,
            "20% (текущая)": 0.20,
            "10% (льготная)": 0.10,
            "Другая...": -1,
        }
        self.selected_rate_str = tk.StringVar(value="22% (прогноз)")
        self.custom_rate_var = tk.StringVar()

        main_container = ttk.Frame(self, padding=10)
        main_container.pack(fill="both", expand=True)
        
        rate_frame = self._create_rate_selector(main_container)
        rate_frame.pack(fill="x", pady=(0, 10), anchor="w")
        
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self.recalculate_on_tab_change)

        self.tab_add = ttk.Frame(self.notebook, padding=10)
        self.tab_extract = ttk.Frame(self.notebook, padding=10)
        self.tab_advance = ttk.Frame(self.notebook, padding=10)
        self.tab_prepayment = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_add, text="НДС начислить")
        self.notebook.add(self.tab_extract, text="НДС выделить")
        self.notebook.add(self.tab_advance, text="Авансы и отгрузка")
        self.notebook.add(self.tab_prepayment, text="Предоплата/Постоплата")

        self.build_tab_add()
        self.build_tab_extract()
        self.build_tab_advance()
        self.build_tab_prepayment()

        self.calc_add_vat()

    def get_current_vat_rate(self) -> float:
        selection = self.selected_rate_str.get()
        rate = self.vat_rates_map.get(selection)
        
        if rate == -1:
            try:
                custom_rate_val = float(self.custom_rate_var.get().replace(",", "."))
                return custom_rate_val / 100.0
            except (ValueError, TypeError):
                return 0.0
        
        return rate if rate is not None else 0.22

    def get_current_vat_share(self) -> float:
        rate = self.get_current_vat_rate()
        if rate <= 0.0: return 0.0
        return rate / (1 + rate)

    def on_rate_or_input_change(self, event=None):
        try:
            tab_index = self.notebook.index(self.notebook.select())
            if tab_index == 0:
                self.calc_add_vat()
            elif tab_index == 1:
                self.calc_extract_vat()
            elif tab_index == 2:
                self.calc_advance()
            elif tab_index == 3:
                self.calc_prepayment()
        except tk.TclError:
            # Возникает при закрытии окна, когда notebook уже не существует
            pass
            
    def recalculate_on_tab_change(self, event=None):
        self.after(50, self.on_rate_or_input_change)

    def on_rate_selection_changed(self, event=None):
        if self.selected_rate_str.get() == "Другая...":
            self.custom_rate_entry.pack(side="left", padx=(5, 0))
            self.custom_rate_label.pack(side="left")
            self.custom_rate_entry.focus()
        else:
            self.custom_rate_entry.pack_forget()
            self.custom_rate_label.pack_forget()
        
        self.on_rate_or_input_change()

    def _create_rate_selector(self, parent):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="Ставка НДС:", font=("Segoe UI", 10, "bold")).pack(side="left")
        rate_combo = ttk.Combobox(
            frame,
            textvariable=self.selected_rate_str,
            values=list(self.vat_rates_map.keys()),
            state="readonly",
            width=15
        )
        rate_combo.pack(side="left", padx=5)
        rate_combo.bind("<<ComboboxSelected>>", self.on_rate_selection_changed)

        self.custom_rate_entry = ttk.Entry(frame, textvariable=self.custom_rate_var, width=6)
        self.custom_rate_entry.bind("<KeyRelease>", self.on_rate_or_input_change)
        self.custom_rate_label = ttk.Label(frame, text="%")
        
        return frame

    def _create_result_row(self, parent, label_text, row):
        text_label = ttk.Label(parent, text=label_text)
        text_label.grid(row=row, column=0, sticky="w", pady=3)
        
        result_label = ttk.Label(parent, text="—", anchor="w")
        result_label.grid(row=row, column=1, sticky="ew", pady=3)
        
        copy_button = ttk.Button(
            parent,
            text="Копировать",
            command=lambda: copy_to_clipboard(self, result_label.cget("text"))
        )
        copy_button.grid(row=row, column=2, padx=10)
        
        return result_label, text_label

    def _create_button_frame(self, parent, clear_command):
        """Создает рамку с кнопками Очистить и Экспорт."""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", pady=10, side="bottom")
        
        ttk.Button(btn_frame, text="Очистить", command=clear_command).pack(side="left")
        ttk.Button(btn_frame, text="Экспорт в Excel", command=self.export_results).pack(side="right")

    # ---------- ВКЛАДКА 1: НАЧИСЛИТЬ НДС ----------
    def build_tab_add(self):
        content_frame = ttk.Frame(self.tab_add)
        content_frame.pack(fill="both", expand=True)

        ttk.Label(content_frame, text="Расчет при добавлении НДС к сумме", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        
        entry_frame = ttk.Frame(content_frame)
        entry_frame.pack(fill="x", pady=10)
        ttk.Label(entry_frame, text="Сумма без НДС (руб):").pack(side="left")
        self.entry_add_base = ttk.Entry(entry_frame)
        self.entry_add_base.pack(side="left", fill="x", expand=True, padx=5)
        self.entry_add_base.insert(0, "100000")
        self.entry_add_base.bind("<KeyRelease>", self.on_rate_or_input_change)

        frame_res = ttk.LabelFrame(content_frame, text="Результат")
        frame_res.pack(fill="both", expand=True, pady=10)
        frame_res.columnconfigure(1, weight=1)

        self.lbl_add_base, _ = self._create_result_row(frame_res, "Сумма без НДС:", 0)
        self.lbl_add_vat, self.lbl_add_vat_text = self._create_result_row(frame_res, "", 1)
        self.lbl_add_total, _ = self._create_result_row(frame_res, "Итого с НДС:", 2)
        self.lbl_add_total.config(font=("Segoe UI", 10, "bold"))
        
        self._create_button_frame(self.tab_add, self.clear_add_tab)

    def calc_add_vat(self):
        try:
            base_str = self.entry_add_base.get().replace(" ", "").replace(",", ".") or "0"
            base = float(base_str)
        except ValueError:
            self.clear_add_tab(clear_input=False)
            return

        rate = self.get_current_vat_rate()
        vat = base * rate
        total = base + vat

        self.lbl_add_base.config(text=format_rub(base))
        self.lbl_add_vat.config(text=format_rub(vat), foreground="green")
        self.lbl_add_total.config(text=format_rub(total))
        self.lbl_add_vat_text.config(text=f"НДС ({format_rate(rate)}%):")

    def clear_add_tab(self, clear_input=True):
        if clear_input:
            self.entry_add_base.delete(0, tk.END)
            self.entry_add_base.insert(0, "0")
        self.lbl_add_base.config(text="—")
        self.lbl_add_vat.config(text="—", foreground="black")
        self.lbl_add_total.config(text="—")
        if clear_input: self.entry_add_base.focus()

    # ---------- ВКЛАДКА 2: ВЫДЕЛИТЬ НДС ----------
    def build_tab_extract(self):
        content_frame = ttk.Frame(self.tab_extract)
        content_frame.pack(fill="both", expand=True)

        ttk.Label(content_frame, text="Выделение НДС из суммы", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        entry_frame = ttk.Frame(content_frame)
        entry_frame.pack(fill="x", pady=10)
        ttk.Label(entry_frame, text="Сумма с НДС (руб):").pack(side="left")
        self.entry_ext_total = ttk.Entry(entry_frame)
        self.entry_ext_total.pack(side="left", fill="x", expand=True, padx=5)
        self.entry_ext_total.insert(0, "122000")
        self.entry_ext_total.bind("<KeyRelease>", self.on_rate_or_input_change)

        frame_res = ttk.LabelFrame(content_frame, text="Результат")
        frame_res.pack(fill="both", expand=True, pady=10)
        frame_res.columnconfigure(1, weight=1)

        self.lbl_ext_total, _ = self._create_result_row(frame_res, "Сумма с НДС:", 0)
        self.lbl_ext_vat, self.lbl_ext_vat_text = self._create_result_row(frame_res, "", 1)
        self.lbl_ext_base, _ = self._create_result_row(frame_res, "Сумма без НДС:", 2)
        self.lbl_ext_base.config(font=("Segoe UI", 10, "bold"))
        
        self._create_button_frame(self.tab_extract, self.clear_extract_tab)
        
    def calc_extract_vat(self):
        try:
            total_str = self.entry_ext_total.get().replace(" ", "").replace(",", ".") or "0"
            total = float(total_str)
        except ValueError:
            self.clear_extract_tab(clear_input=False)
            return

        vat_share = self.get_current_vat_share()
        vat = total * vat_share
        base = total - vat

        self.lbl_ext_total.config(text=format_rub(total))
        self.lbl_ext_vat.config(text=format_rub(vat), foreground="green")
        self.lbl_ext_base.config(text=format_rub(base))
        rate = self.get_current_vat_rate()
        self.lbl_ext_vat_text.config(text=f"в т.ч. НДС ({format_rate(rate)}%):")

    def clear_extract_tab(self, clear_input=True):
        if clear_input:
            self.entry_ext_total.delete(0, tk.END)
            self.entry_ext_total.insert(0, "0")
        self.lbl_ext_total.config(text="—")
        self.lbl_ext_vat.config(text="—", foreground="black")
        self.lbl_ext_base.config(text="—")
        if clear_input: self.entry_ext_total.focus()

    # ---------- ВКЛАДКА 3: АВАНСЫ ----------
    def build_tab_advance(self):
        self.tab_advance.columnconfigure([0, 1], weight=1)
        self.tab_advance.rowconfigure(3, weight=1) # Main frames row

        ttk.Label(self.tab_advance, text="Расчет НДС по авансам и отгрузке", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        ttk.Label(self.tab_advance, text="Аванс с НДС (руб):").grid(row=1, column=0, sticky="w", padx=5)
        self.entry_adv_total = ttk.Entry(self.tab_advance)
        self.entry_adv_total.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        self.entry_adv_total.insert(0, "122000")
        self.entry_adv_total.bind("<KeyRelease>", self.on_rate_or_input_change)

        ttk.Label(self.tab_advance, text="Отгрузка без НДС (руб):").grid(row=1, column=1, sticky="w", padx=5)
        self.entry_adv_rest = ttk.Entry(self.tab_advance)
        self.entry_adv_rest.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.entry_adv_rest.insert(0, "100000")
        self.entry_adv_rest.bind("<KeyRelease>", self.on_rate_or_input_change)
        
        frame_adv = ttk.LabelFrame(self.tab_advance, text="Расчет по авансу")
        frame_adv.grid(row=3, column=0, sticky="nsew", padx=5, pady=10)
        frame_adv.columnconfigure(1, weight=1)

        self.lbl_adv_total, _ = self._create_result_row(frame_adv, "Аванс с НДС:", 0)
        self.lbl_adv_vat, self.lbl_adv_vat_text = self._create_result_row(frame_adv, "НДС с аванса:", 1)
        self.lbl_adv_base, _ = self._create_result_row(frame_adv, "База аванса:", 2)

        frame_rest = ttk.LabelFrame(self.tab_advance, text="Расчет по отгрузке")
        frame_rest.grid(row=3, column=1, sticky="nsew", padx=5, pady=10)
        frame_rest.columnconfigure(1, weight=1)
        
        self.lbl_rest_base, _ = self._create_result_row(frame_rest, "Отгрузка без НДС:", 0)
        self.lbl_rest_vat, self.lbl_rest_vat_text = self._create_result_row(frame_rest, "НДС с отгрузки:", 1)
        self.lbl_rest_total, _ = self._create_result_row(frame_rest, "Сумма отгрузки с НДС:", 2)

        frame_total = ttk.LabelFrame(self.tab_advance, text="Итого НДС к уплате")
        frame_total.grid(row=4, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        frame_total.columnconfigure(0, weight=1)

        self.lbl_total_vat = ttk.Label(frame_total, text="Всего НДС (аванс + отгрузка): —", font=("Segoe UI", 10, "bold"))
        self.lbl_total_vat.grid(row=0, column=0, sticky="w", pady=5, padx=5)
        
        copy_button = ttk.Button(frame_total, text="Копировать итог НДС", command=lambda: copy_to_clipboard(self, self.lbl_total_vat.cget("text").split(":")[-1]))
        copy_button.grid(row=0, column=1, padx=5, pady=5, sticky="e")
        
        btn_frame = ttk.Frame(self.tab_advance)
        btn_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=5, pady=10)
        ttk.Button(btn_frame, text="Очистить всё", command=self.clear_advance_tab).pack(side="left")
        ttk.Button(btn_frame, text="Экспорт в Excel", command=self.export_results).pack(side="right")

    def calc_advance(self):
        try:
            adv_str = self.entry_adv_total.get().replace(" ", "").replace(",", ".") or "0"
            rest_str = self.entry_adv_rest.get().replace(" ", "").replace(",", ".") or "0"
            advance_total = float(adv_str)
            rest_base = float(rest_str)
        except ValueError:
            self.clear_advance_tab(clear_input=False)
            return

        rate = self.get_current_vat_rate()
        rate_str = format_rate(rate)
        
        adv_vat_share = self.get_current_vat_share()
        adv_vat = advance_total * adv_vat_share
        adv_base = advance_total - adv_vat

        self.lbl_adv_total.config(text=format_rub(advance_total))
        self.lbl_adv_vat.config(text=format_rub(adv_vat), foreground="blue")
        self.lbl_adv_base.config(text=format_rub(adv_base))

        rest_vat = rest_base * rate
        rest_total = rest_base + rest_vat

        self.lbl_rest_base.config(text=format_rub(rest_base))
        self.lbl_rest_vat.config(text=format_rub(rest_vat), foreground="blue")
        self.lbl_rest_total.config(text=format_rub(rest_total))
        
        self.lbl_adv_vat_text.config(text=f"НДС с аванса ({rate_str}%):")
        self.lbl_rest_vat_text.config(text=f"НДС с отгрузки ({rate_str}%):")

        total_vat = adv_vat + rest_vat
        self.lbl_total_vat.config(
            text=f"Всего НДС (аванс + отгрузка): {format_rub(total_vat)}",
            foreground="purple"
        )
        
    def clear_advance_tab(self, clear_input=True):
        if clear_input:
            self.entry_adv_total.delete(0, tk.END)
            self.entry_adv_total.insert(0, "0")
            self.entry_adv_rest.delete(0, tk.END)
            self.entry_adv_rest.insert(0, "0")
            self.entry_adv_total.focus()
            
        self.lbl_adv_total.config(text="—")
        self.lbl_adv_vat.config(text="—", foreground="black")
        self.lbl_adv_base.config(text="—")
        self.lbl_rest_base.config(text="—")
        self.lbl_rest_vat.config(text="—", foreground="black")
        self.lbl_rest_total.config(text="—")
        self.lbl_total_vat.config(text="Всего НДС (аванс + отгрузка): —", foreground="black")

    # ---------- ВКЛАДКА 4: ПРЕДОПЛАТА/ПОСТООПЛАТА ----------
    def build_tab_prepayment(self):
        # --- Структура вкладки ---
        self.tab_prepayment.columnconfigure(0, weight=1)
        
        # --- Фрейм для полей ввода ---
        input_frame = ttk.Frame(self.tab_prepayment)
        input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text="Общая стоимость работ (без НДС):").grid(row=0, column=0, sticky="w", padx=5)
        self.entry_prep_total_cost = ttk.Entry(input_frame)
        self.entry_prep_total_cost.grid(row=0, column=1, sticky="ew", padx=5, pady=(0, 5))
        self.entry_prep_total_cost.insert(0, "500000")
        self.entry_prep_total_cost.bind("<KeyRelease>", self.on_rate_or_input_change)

        ttk.Label(input_frame, text="Процент предоплаты (%):").grid(row=1, column=0, sticky="w", padx=5)
        self.entry_prep_percentage = ttk.Entry(input_frame)
        self.entry_prep_percentage.grid(row=1, column=1, sticky="ew", padx=5)
        self.entry_prep_percentage.insert(0, "30")
        self.entry_prep_percentage.bind("<KeyRelease>", self.on_rate_or_input_change)

        # --- Фреймы для результатов ---
        results_frame = ttk.Frame(self.tab_prepayment)
        results_frame.grid(row=1, column=0, sticky="nsew")
        results_frame.columnconfigure([0, 1], weight=1)
        self.tab_prepayment.rowconfigure(1, weight=1)

        # --- Блок Предоплаты ---
        frame_prepayment = ttk.LabelFrame(results_frame, text="Расчет Предоплаты")
        frame_prepayment.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=5)
        frame_prepayment.columnconfigure(1, weight=1)
        
        self.lbl_prep_base, _ = self._create_result_row(frame_prepayment, "Сумма предоплаты (без НДС):", 0)
        self.lbl_prep_vat, self.lbl_prep_vat_text = self._create_result_row(frame_prepayment, "НДС с предоплаты:", 1)
        self.lbl_prep_total, _ = self._create_result_row(frame_prepayment, "Итого предоплата (с НДС):", 2)
        self.lbl_prep_total.config(font=("Segoe UI", 10, "bold"))

        # --- Блок Постоплаты ---
        frame_postpayment = ttk.LabelFrame(results_frame, text="Расчет Постоплаты")
        frame_postpayment.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)
        frame_postpayment.columnconfigure(1, weight=1)

        self.lbl_post_base, _ = self._create_result_row(frame_postpayment, "Сумма постоплаты (без НДС):", 0)
        self.lbl_post_vat, self.lbl_post_vat_text = self._create_result_row(frame_postpayment, "НДС с постоплаты:", 1)
        self.lbl_post_total, _ = self._create_result_row(frame_postpayment, "Итого постоплата (с НДС):", 2)
        self.lbl_post_total.config(font=("Segoe UI", 10, "bold"))

        # --- Итоговый блок ---
        frame_final = ttk.LabelFrame(self.tab_prepayment, text="Итоги по проекту")
        frame_final.grid(row=2, column=0, sticky="ew", pady=5)
        frame_final.columnconfigure(1, weight=1)

        self.lbl_final_base, _ = self._create_result_row(frame_final, "Общая стоимость (без НДС):", 0)
        self.lbl_final_vat, _ = self._create_result_row(frame_final, "Общий НДС по проекту:", 1)
        self.lbl_final_total, _ = self._create_result_row(frame_final, "Общая стоимость (с НДС):", 2)
        self.lbl_final_total.config(font=("Segoe UI", 11, "bold"))
        
        # --- Кнопки ---
        btn_frame = ttk.Frame(self.tab_prepayment)
        btn_frame.grid(row=3, column=0, sticky="ew", pady=(5, 0))
        ttk.Button(btn_frame, text="Очистить", command=self.clear_prepayment_tab).pack(side="left")
        ttk.Button(btn_frame, text="Экспорт в Excel", command=self.export_results).pack(side="right")
        
    def calc_prepayment(self):
        try:
            total_cost_str = self.entry_prep_total_cost.get().replace(" ", "").replace(",", ".") or "0"
            percentage_str = self.entry_prep_percentage.get().replace(" ", "").replace(",", ".") or "0"
            
            total_cost_base = float(total_cost_str)
            prepayment_percent = float(percentage_str)
            
            if not (0 <= prepayment_percent <= 100):
                self.clear_prepayment_tab(clear_input=False)
                messagebox.showwarning("Неверное значение", "Процент предоплаты должен быть в диапазоне от 0 до 100.")
                return

        except ValueError:
            self.clear_prepayment_tab(clear_input=False)
            return
            
        rate = self.get_current_vat_rate()
        rate_str = format_rate(rate)
        
        # Расчеты
        prepayment_base = total_cost_base * (prepayment_percent / 100.0)
        prepayment_vat = prepayment_base * rate
        prepayment_total = prepayment_base + prepayment_vat
        
        postpayment_base = total_cost_base - prepayment_base
        postpayment_vat = postpayment_base * rate
        postpayment_total = postpayment_base + postpayment_vat
        
        total_vat = total_cost_base * rate
        total_cost_with_vat = total_cost_base + total_vat
        
        # Обновление лейблов предоплаты
        self.lbl_prep_base.config(text=format_rub(prepayment_base))
        self.lbl_prep_vat.config(text=format_rub(prepayment_vat), foreground="green")
        self.lbl_prep_total.config(text=format_rub(prepayment_total))
        self.lbl_prep_vat_text.config(text=f"НДС с предоплаты ({rate_str}%):")
        
        # Обновление лейблов постоплаты
        self.lbl_post_base.config(text=format_rub(postpayment_base))
        self.lbl_post_vat.config(text=format_rub(postpayment_vat), foreground="green")
        self.lbl_post_total.config(text=format_rub(postpayment_total))
        self.lbl_post_vat_text.config(text=f"НДС с постоплаты ({rate_str}%):")
        
        # Обновление итоговых лейблов
        self.lbl_final_base.config(text=format_rub(total_cost_base))
        self.lbl_final_vat.config(text=format_rub(total_vat), foreground="purple")
        self.lbl_final_total.config(text=format_rub(total_cost_with_vat))

    def clear_prepayment_tab(self, clear_input=True):
        if clear_input:
            self.entry_prep_total_cost.delete(0, tk.END)
            self.entry_prep_total_cost.insert(0, "0")
            self.entry_prep_percentage.delete(0, tk.END)
            self.entry_prep_percentage.insert(0, "30")
            self.entry_prep_total_cost.focus()

        labels = [
            self.lbl_prep_base, self.lbl_prep_vat, self.lbl_prep_total,
            self.lbl_post_base, self.lbl_post_vat, self.lbl_post_total,
            self.lbl_final_base, self.lbl_final_vat, self.lbl_final_total
        ]
        for lbl in labels:
            lbl.config(text="—", foreground="black") # Сброс цвета для НДС
        self.lbl_final_total.config(foreground="black")
        self.lbl_final_vat.config(foreground="black")


    def export_results(self):
        """
        Экспортирует текущие расчеты в файл Excel с формулами.
        """
        tab_index = self.notebook.index(self.notebook.select())
        rate = self.get_current_vat_rate()
        rate_str = format_rate(rate)

        # Вспомогательная функция для безопасного получения числового значения из поля ввода
        def _parse_input(entry_widget) -> float:
            try:
                return float(entry_widget.get().replace(" ", "").replace(",", ".") or "0")
            except (ValueError, TypeError):
                return 0.0

        # --- Общие настройки для Excel ---
        wb = Workbook()
        ws = wb.active
        ws.title = "Расчет НДС"
        
        # Стили
        header_font = Font(bold=True)
        category_font = Font(bold=True, color="FF4F81BD")
        category_alignment = Alignment(horizontal="center")
        money_format = "#,##0.00 ₽"
        bold_font = Font(bold=True)
        
        # Заголовки
        ws.append(["Параметр", "Значение"])
        ws['A1'].font = header_font
        ws['B1'].font = header_font
        ws['D1'] = rate
        ws['D1'].number_format = '0.00%'
        ws['D1'].font = header_font
        ws.cell(row=1, column=3, value="Ставка НДС:").font = header_font
        
        default_filename = ""
        
        # --- Логика для каждой вкладки ---
        if tab_index == 0:  # НАЧИСЛИТЬ НДС
            default_filename = f"Расчет_начисление_НДС_{rate_str}%.xlsx"
            base_amount = _parse_input(self.entry_add_base)
            
            ws.append(["Сумма без НДС", base_amount])
            ws.append([f"НДС ({rate_str}%)", "=B2*D1"])
            ws.append(["Итого с НДС", "=B2+B3"])
            
            # Применение стилей и форматов
            for row_num in [2, 3, 4]:
                ws[f'B{row_num}'].number_format = money_format
            ws['A4'].font = ws['B4'].font = bold_font

        elif tab_index == 1:  # ВЫДЕЛИТЬ НДС
            default_filename = f"Расчет_выделение_НДС_{rate_str}%.xlsx"
            total_amount = _parse_input(self.entry_ext_total)
            
            ws.append(["Сумма с НДС", total_amount])
            ws.append([f"в т.ч. НДС ({rate_str}%)", "=B2*(D1/(1+D1))"])
            ws.append(["Сумма без НДС", "=B2-B3"])

            # Применение стилей и форматов
            for row_num in [2, 3, 4]:
                ws[f'B{row_num}'].number_format = money_format
            ws['A4'].font = ws['B4'].font = bold_font

        elif tab_index == 2:  # АВАНСЫ И ОТГРУЗКА
            default_filename = f"Расчет_аванс_и_отгрузка_{rate_str}%.xlsx"
            advance_total = _parse_input(self.entry_adv_total)
            rest_base = _parse_input(self.entry_adv_rest)

            # --- Расчет по авансу ---
            ws.append(["--- РАСЧЕТ ПО АВАНСУ ---", None])
            ws.merge_cells('A2:B2')
            ws['A2'].font = category_font
            ws['A2'].alignment = category_alignment
            
            ws.append(["Аванс с НДС:", advance_total])
            ws.append([f"НДС с аванса ({rate_str}%):", "=B3*($D$1/(1+$D$1))"])
            ws.append(["База аванса:", "=B3-B4"])
            
            # --- Расчет по отгрузке ---
            ws.append([None, None]) # Пустая строка
            ws.append(["--- РАСЧЕТ ПО ОТГРУЗКЕ ---", None])
            ws.merge_cells('A7:B7')
            ws['A7'].font = category_font
            ws['A7'].alignment = category_alignment
            
            ws.append(["Отгрузка без НДС:", rest_base])
            ws.append([f"НДС с отгрузки ({rate_str}%):", "=B8*$D$1"])
            ws.append(["Сумма отгрузки с НДС:", "=B8+B9"])
            
            # --- Итого ---
            ws.append([None, None]) # Пустая строка
            ws.append(["--- ИТОГО ПО ДОГОВОРУ ---", None])
            ws.merge_cells('A12:B12')
            ws['A12'].font = category_font
            ws['A12'].alignment = category_alignment
            
            ws.append(["Всего НДС (аванс + отгрузка):", "=B4+B9"])
            ws['A13'].font = ws['B13'].font = bold_font

            # Применение денежного формата
            for row_num in [3, 4, 5, 8, 9, 10, 13]:
                ws[f'B{row_num}'].number_format = money_format

        elif tab_index == 3: # ПРЕДОПЛАТА/ПОСТООПЛАТА
            default_filename = f"Расчет_предоплата_{rate_str}%.xlsx"
            total_cost = _parse_input(self.entry_prep_total_cost)
            percentage = _parse_input(self.entry_prep_percentage) / 100.0

            # Доп. ячейка для процента предоплаты
            ws.cell(row=2, column=3, value="Процент предоплаты:").font = header_font
            ws['D2'] = percentage
            ws['D2'].number_format = '0.00%'

            ws.append(["Общая стоимость работ (без НДС)", total_cost])
            ws.append([None, None])

            # Предоплата
            ws.append(["--- ПРЕДОПЛАТА ---", None]); ws.merge_cells('A5:B5'); ws['A5'].font = category_font; ws['A5'].alignment = category_alignment
            ws.append(["Сумма предоплаты (без НДС)", "=B3*D2"])
            ws.append([f"НДС с предоплаты ({rate_str}%)", "=B6*D1"])
            ws.append(["Итого предоплата (с НДС)", "=B6+B7"]); ws['A8'].font = ws['B8'].font = bold_font
            ws.append([None, None])
            
            # Постоплата
            ws.append(["--- ПОСТОПЛАТА ---", None]); ws.merge_cells('A10:B10'); ws['A10'].font = category_font; ws['A10'].alignment = category_alignment
            ws.append(["Сумма постоплаты (без НДС)", "=B3-B6"])
            ws.append([f"НДС с постоплаты ({rate_str}%)", "=B11*D1"])
            ws.append(["Итого постоплата (с НДС)", "=B11+B12"]); ws['A13'].font = ws['B13'].font = bold_font
            ws.append([None, None])

            # Итоги
            ws.append(["--- ИТОГИ ПО ПРОЕКТУ ---", None]); ws.merge_cells('A15:B15'); ws['A15'].font = category_font; ws['A15'].alignment = category_alignment
            ws.append(["Общая стоимость (без НДС)", "=B6+B11"])
            ws.append(["Общий НДС по проекту", "=B7+B12"])
            ws.append(["Общая стоимость (с НДС)", "=B16+B17"]); ws['A18'].font = ws['B18'].font = bold_font

            # Форматирование
            for row_num in [3, 6, 7, 8, 11, 12, 13, 16, 17, 18]:
                ws[f'B{row_num}'].number_format = money_format
                
        else:
            return  # Неизвестная вкладка

        # --- Сохранение файла ---
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel-файл", "*.xlsx"), ("Все файлы", "*.*")],
            initialfile=default_filename,
            title="Сохранить отчет в Excel"
        )

        if not filepath:
            return

        try:
            # Автоподбор ширины колонок
            for col_idx, column_cells in enumerate(ws.columns, 1):
                max_length = 0
                column_letter = get_column_letter(col_idx)
                for cell in column_cells:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 3)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            ws.column_dimensions['B'].width = 25

            wb.save(filepath)
            messagebox.showinfo("Экспорт успешно завершен", f"Данные сохранены в файл:\n{filepath}")

        except Exception as e:
            messagebox.showerror("Ошибка экспорта", f"Не удалось сохранить файл.\nОшибка: {e}")


if __name__ == "__main__":
    app = NDSProApp()
    app.mainloop()
