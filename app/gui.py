import sys
import subprocess
import threading
import queue
from pathlib import Path
from tkinter import Tk, StringVar, BooleanVar, BOTH, END, LEFT, RIGHT, X, Y, filedialog, messagebox
from tkinter import ttk

import pandas as pd

from config import DEFAULT_PROJECT_DIR, DEFAULT_REGISTERS_DIR, DEFAULT_OUTPUT_WORKBOOK_NAME
from core.paths import ProjectPaths, ensure_project_structure, scan_register_pdfs, _norm_folder_name
from core.excel_format import write_multi_sheet_excel_openpyxl
from registers.our_registers import generate_sales_register, generate_purchase_register, build_sales_summary_df, build_purchase_summary_df
from compare.register_compare import create_pdf_comparison_workbook

def run_app() -> None:
    root = Tk()
    RejestrApp(root)
    root.mainloop()

class RejestrApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("ISS-Chem бухгалтерський helper — Реєстри")
        self.root.geometry("980x680")

        self.project_dir = StringVar(value=str(DEFAULT_PROJECT_DIR))
        self.output_xlsx = StringVar(value="")
        self.generate_sales_var = BooleanVar(value=True)
        self.generate_purchase_var = BooleanVar(value=True)
        self.scan_pdf_var = BooleanVar(value=True)
        self.log_queue: queue.Queue[str] = queue.Queue()

        self._build_ui()
        self._on_project_changed()
        self._poll_log_queue()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        main = ttk.Frame(self.root)
        main.pack(fill=BOTH, expand=True)

        paths_box = ttk.LabelFrame(main, text="Шляхи")
        paths_box.pack(fill=X, **pad)

        self._path_row(paths_box, "Папка проєкту", self.project_dir, self.choose_project_dir, 0)
        self._path_row(paths_box, "Файл результату", self.output_xlsx, self.choose_output_xlsx, 1)

        options_box = ttk.LabelFrame(main, text="Що зробити")
        options_box.pack(fill=X, **pad)
        ttk.Checkbutton(options_box, text="Згенерувати Реєстр продажу", variable=self.generate_sales_var).pack(side=LEFT, padx=10, pady=8)
        ttk.Checkbutton(options_box, text="Згенерувати Реєстр закупівлі", variable=self.generate_purchase_var).pack(side=LEFT, padx=10, pady=8)
        ttk.Checkbutton(options_box, text="Додати список PDF-реєстрів з папки ‘Реєстри’", variable=self.scan_pdf_var).pack(side=LEFT, padx=10, pady=8)

        actions = ttk.Frame(main)
        actions.pack(fill=X, **pad)
        ttk.Button(actions, text="1. Перевірити структуру папок", command=self.create_structure).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="2. Запустити генерацію", command=self.run_generation_threaded).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="PDF → Excel + порівняння…", command=self.compare_one_pdf_threaded).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="Порівняти всі PDF", command=self.compare_all_pdfs_threaded).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="Відкрити папку Реєстри", command=self.open_registers_dir).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="Очистити лог", command=lambda: self.log_text.delete("1.0", END)).pack(side=RIGHT, padx=5)

        log_box = ttk.LabelFrame(main, text="Лог")
        log_box.pack(fill=BOTH, expand=True, **pad)
        self.log_text = __import__("tkinter").Text(log_box, wrap="word", height=22)
        scrollbar = ttk.Scrollbar(log_box, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        hint = (
            "Назви PDF підтримуються так: ‘Реєстр Закупівлі 012026-032026.pdf’ "
            "або ‘Реєстр Закупівлі 01012026-15032026.pdf’. "
            "MMYYYY означає весь місяць, DDMMYYYY — точну дату."
        )
        ttk.Label(main, text=hint, foreground="#555").pack(fill=X, padx=14, pady=(0, 8))

    def _path_row(self, parent, label: str, var: StringVar, command, row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=5)
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
        ttk.Button(parent, text="Обрати…", command=command).grid(row=row, column=2, padx=8, pady=5)
        parent.columnconfigure(1, weight=1)

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

    def _on_project_changed(self) -> None:
        project = Path(self.project_dir.get()).expanduser()
        paths = ProjectPaths(project)
        self.output_xlsx.set(str(paths.output_registers_workbook))
    def choose_project_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=str(Path.home() / "Desktop"), title="Оберіть папку ISS-Chem sp. z o. o.")
        if selected:
            selected_path = Path(selected).expanduser()
            # If the user accidentally chooses the Реєстри folder itself, use its parent
            # as the project folder. Otherwise the app would create Реєстри/Реєстри.
            if _norm_folder_name(selected_path.name) == _norm_folder_name(DEFAULT_REGISTERS_DIR):
                selected_path = selected_path.parent
            self.project_dir.set(str(selected_path))
            self._on_project_changed()

    def choose_output_xlsx(self) -> None:
        selected = filedialog.asksaveasfilename(
            initialdir=str(ProjectPaths(Path(self.project_dir.get())).registers_dir),
            initialfile=DEFAULT_OUTPUT_WORKBOOK_NAME,
            defaultextension=".xlsx",
            title="Куди зберегти файл реєстрів",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self.output_xlsx.set(selected)

    def create_structure(self) -> None:
        try:
            paths = ensure_project_structure(Path(self.project_dir.get()))
            self.output_xlsx.set(str(paths.output_registers_workbook))
            self.log(f"Структуру папок перевірено: {paths.project_dir}")
            self.log(f"Реєстри: {paths.registers_dir}")
            self.log(f"Прибуток: {paths.profit_dir}")
            self.log(f"Бухгалтерський звіт: {paths.accounting_report_dir}")
        except Exception as exc:
            messagebox.showerror("Помилка", str(exc))
            self.log(f"ПОМИЛКА: {exc}")

    def run_generation_threaded(self) -> None:
        thread = threading.Thread(target=self.run_generation, daemon=True)
        thread.start()

    def run_generation(self) -> None:
        try:
            paths = ensure_project_structure(Path(self.project_dir.get()))
            source = paths.accounting_table.expanduser().resolve()
            output = Path(self.output_xlsx.get()).expanduser().resolve()

            if not source.exists():
                raise FileNotFoundError(f"Не знайдено accounting_table.xlsx: {source}")
            if not (self.generate_sales_var.get() or self.generate_purchase_var.get() or self.scan_pdf_var.get()):
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

            if self.scan_pdf_var.get():
                self.log("Сканую PDF-реєстри у папці ‘Реєстри’…")
                pdf_index = scan_register_pdfs(paths.registers_dir)
                sheets["PDF rejestry"] = pdf_index
                self.log(f"Знайдено PDF-реєстрів з коректною назвою: {len(pdf_index)}.")

            write_multi_sheet_excel_openpyxl(output, sheets)
            self.log(f"ГОТОВО. Файл збережено: {output}")
            messagebox.showinfo("Готово", f"Файл створено:\n{output}")
        except Exception as exc:
            self.log(f"ПОМИЛКА: {exc}")
            messagebox.showerror("Помилка", str(exc))


    def compare_one_pdf_threaded(self) -> None:
        try:
            paths = ensure_project_structure(Path(self.project_dir.get()))
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
            paths = ensure_project_structure(Path(self.project_dir.get()))
            if not paths.accounting_table.exists():
                raise FileNotFoundError(f"Не знайдено accounting_table.xlsx: {paths.accounting_table}")
            self.log(f"Створюю порівняння для PDF: {pdf_path.name}")
            output = create_pdf_comparison_workbook(paths, pdf_path, log_func=self.log)
            self.log(f"ГОТОВО. Файл порівняння створено: {output}")
            messagebox.showinfo("Готово", f"Файл порівняння створено:\n{output}")
        except Exception as exc:
            self.log(f"ПОМИЛКА PDF-порівняння: {exc}")
            messagebox.showerror("Помилка", str(exc))

    def compare_all_pdfs_threaded(self) -> None:
        thread = threading.Thread(target=self.compare_all_pdfs, daemon=True)
        thread.start()

    def compare_all_pdfs(self) -> None:
        try:
            paths = ensure_project_structure(Path(self.project_dir.get()))
            if not paths.accounting_table.exists():
                raise FileNotFoundError(f"Не знайдено accounting_table.xlsx: {paths.accounting_table}")
            pdf_index = scan_register_pdfs(paths.registers_dir)
            if pdf_index.empty:
                raise ValueError(f"Не знайшов PDF з коректною назвою у папці: {paths.registers_dir}")
            outputs = []
            for _, row in pdf_index.iterrows():
                pdf_path = Path(row["Path"])
                self.log(f"Порівнюю PDF: {pdf_path.name}")
                outputs.append(create_pdf_comparison_workbook(paths, pdf_path, log_func=self.log))
            self.log(f"ГОТОВО. Створено файлів порівняння: {len(outputs)}")
            messagebox.showinfo("Готово", "Створено файлів порівняння: " + str(len(outputs)))
        except Exception as exc:
            self.log(f"ПОМИЛКА PDF-порівняння: {exc}")
            messagebox.showerror("Помилка", str(exc))

    def open_registers_dir(self) -> None:
        try:
            paths = ensure_project_structure(Path(self.project_dir.get()))
            folder = str(paths.registers_dir)
            if sys.platform == "darwin":
                subprocess.run(["open", folder], check=False)
            elif sys.platform.startswith("win"):
                subprocess.run(["explorer", folder], check=False)
            else:
                subprocess.run(["xdg-open", folder], check=False)
        except Exception as exc:
            self.log(f"Не вдалося відкрити папку: {exc}")
