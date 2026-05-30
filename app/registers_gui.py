import threading
import queue
from datetime import datetime
from pathlib import Path
from tkinter import BooleanVar, BOTH, END, LEFT, RIGHT, X, Y, filedialog, messagebox
from tkinter import ttk

import pandas as pd

from config import DEFAULT_PROJECT_DIR
from core.paths import ensure_project_structure
from core.excel_format import write_multi_sheet_excel_openpyxl
from registers.our_registers import generate_sales_register, generate_purchase_register, build_sales_summary_df, build_purchase_summary_df

class RejestrApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("ISS-Chem бухгалтерський helper — Реєстри")
        self.root.geometry("820x560")

        self.generate_sales_var = BooleanVar(value=True)
        self.generate_purchase_var = BooleanVar(value=True)
        self.log_queue: queue.Queue[str] = queue.Queue()

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        main = ttk.Frame(self.root)
        main.pack(fill=BOTH, expand=True)

        options_box = ttk.LabelFrame(main, text="Що зробити")
        options_box.pack(fill=X, **pad)
        ttk.Checkbutton(options_box, text="Згенерувати Реєстр продажу", variable=self.generate_sales_var).pack(side=LEFT, padx=10, pady=8)
        ttk.Checkbutton(options_box, text="Згенерувати Реєстр закупівлі", variable=self.generate_purchase_var).pack(side=LEFT, padx=10, pady=8)

        actions = ttk.Frame(main)
        actions.pack(fill=X, **pad)
        ttk.Button(actions, text="Запустити генерацію", command=self.run_generation_threaded).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="PDF → Excel + порівняння…", command=self.compare_one_pdf_threaded).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="Очистити лог", command=lambda: self.log_text.delete("1.0", END)).pack(side=RIGHT, padx=5)

        log_box = ttk.LabelFrame(main, text="Лог")
        log_box.pack(fill=BOTH, expand=True, **pad)
        self.log_text = __import__("tkinter").Text(log_box, wrap="word", height=22)
        scrollbar = ttk.Scrollbar(log_box, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

    def log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] {message}\n")

    def _poll_log_queue(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert(END, msg)
                self.log_text.see(END)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log_queue)

    def run_generation_threaded(self) -> None:
        thread = threading.Thread(target=self.run_generation, daemon=True)
        thread.start()

    def run_generation(self) -> None:
        try:
            paths = ensure_project_structure(DEFAULT_PROJECT_DIR)
            source = paths.accounting_table.expanduser().resolve()
            output = paths.output_registers_workbook.expanduser().resolve()

            if not source.exists():
                raise FileNotFoundError(f"Не знайдено accounting_table.xlsx: {source}")
            if not (self.generate_sales_var.get() or self.generate_purchase_var.get()):
                raise ValueError("Оберіть хоча б одну дію.")

            self.log(f"Старт. Дані беру автоматично з: {source}")
            sheets: dict[str, pd.DataFrame] = {}

            if self.generate_sales_var.get():
                self.log("Генерую Реєстр продажу…")
                sales = generate_sales_register(source)
                sheets["Rejestr sprzedazy"] = sales
                sheets["Summary sprzedazy"] = build_sales_summary_df(sales)
                self.log(f"Реєстр продажу готовий: {len(sales)} рядків.")

            if self.generate_purchase_var.get():
                self.log("Генерую Реєстр закупівлі…")
                purchase = generate_purchase_register(source)
                sheets["Rejestr zakupu"] = purchase
                sheets["Summary zakupu"] = build_purchase_summary_df(purchase)
                problems = purchase[purchase["Status"].ne("OK") | purchase["Warning"].fillna("").ne("")]
                if not problems.empty:
                    sheets["Problems zakupu"] = problems
                    self.log(f"У Реєстрі закупівлі знайдено проблемні рядки: {len(problems)}. Додано лист Problems zakupu.")
                self.log(f"Реєстр закупівлі готовий: {len(purchase)} рядків.")

            write_multi_sheet_excel_openpyxl(output, sheets)
            self.log(f"ГОТОВО. Файл збережено: {output}")
            messagebox.showinfo("Готово", f"Файл створено:\n{output}")
        except Exception as exc:
            self.log(f"ПОМИЛКА: {exc}")
            messagebox.showerror("Помилка", str(exc))

    def compare_one_pdf_threaded(self) -> None:
        try:
            paths = ensure_project_structure(DEFAULT_PROJECT_DIR)
            selected = filedialog.askopenfilename(
                initialdir=str(paths.registers_dir),
                title="Оберіть PDF-реєстр від бухгалтерії",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            )
            if not selected:
                return
            thread = threading.Thread(target=self.compare_one_pdf, args=(Path(selected),), daemon=True)
            thread.start()
        except Exception as exc:
            self.log(f"ПОМИЛКА: {exc}")
            messagebox.showerror("Помилка", str(exc))

    def compare_one_pdf(self, pdf_path: Path) -> None:
        try:
            paths = ensure_project_structure(DEFAULT_PROJECT_DIR)
            if not paths.accounting_table.exists():
                raise FileNotFoundError(f"Не знайдено accounting_table.xlsx: {paths.accounting_table}")
            self.log(f"Створюю порівняння для PDF: {pdf_path.name}")
            from compare.register_compare import create_pdf_comparison_workbook
            output = create_pdf_comparison_workbook(paths, pdf_path, log_func=self.log)
            self.log(f"ГОТОВО. Файл порівняння створено: {output}")
            messagebox.showinfo("Готово", f"Файл порівняння створено:\n{output}")
        except Exception as exc:
            self.log(f"ПОМИЛКА PDF-порівняння: {exc}")
            messagebox.showerror("Помилка", str(exc))
