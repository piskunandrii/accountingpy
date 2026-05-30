import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import pandas as pd

from config import DEFAULT_REGISTERS_DIR, DEFAULT_PROFIT_DIR, DEFAULT_ACCOUNTING_REPORT_DIR, DEFAULT_OUTPUT_WORKBOOK_NAME
from core.dates import parse_compact_register_date

PDF_REGISTER_RE = re.compile(
    r"^Реєстр\s+(?P<kind>Закупівлі|Закупівель|Продажу|Продажів)\s+"
    r"(?P<start>\d{4}|\d{6}|\d{8})-(?P<end>\d{4}|\d{6}|\d{8})\.pdf$",
    flags=re.IGNORECASE,
)

def _norm_folder_name(value: str) -> str:
    """Normalize folder names so macOS Unicode variants compare correctly."""
    return unicodedata.normalize("NFC", value).casefold().strip()

def find_existing_registers_dir(project_dir: Path) -> Path:
    """
    Return ONLY an already-existing 'Реєстри' folder.

    This function intentionally does NOT create the folder. If the app cannot
    find the existing registers folder, it raises an error with the exact path
    it checked. This prevents accidental parallel folders such as Desktop/Реєстри
    or Реєстри/Реєстри.
    """
    project_dir = project_dir.expanduser().resolve()
    wanted = _norm_folder_name(DEFAULT_REGISTERS_DIR
                               )

    # If the selected folder is already 'Реєстри', use it directly.
    if _norm_folder_name(project_dir.name) == wanted:
        if project_dir.is_dir():
            return project_dir
        raise FileNotFoundError(f"Папка Реєстри не існує: {project_dir}")

    # Prefer an existing direct child whose normalized name is 'Реєстри'.
    if project_dir.exists():
        for child in project_dir.iterdir():
            if child.is_dir() and _norm_folder_name(child.name) == wanted:
                return child

    expected = project_dir / DEFAULT_REGISTERS_DIR
    raise FileNotFoundError(
        "Не знайшов існуючу папку 'Реєстри'.\n"
        f"Перевірений шлях: {expected}\n"
        "У полі 'Папка проєкту' має бути папка ISS-Chem sp. z o. o., "
        "а не Desktop, AccountingPro або інша папка."
    )

def find_existing_child_dir(project_dir: Path, folder_name: str) -> Path:
    project_dir = project_dir.expanduser().resolve()
    wanted = _norm_folder_name(folder_name)
    if project_dir.exists():
        for child in project_dir.iterdir():
            if child.is_dir() and _norm_folder_name(child.name) == wanted:
                return child
    return project_dir / folder_name

@dataclass
class ProjectPaths:
    project_dir: Path
    _registers_dir: Optional[Path] = None

    @property
    def registers_dir(self) -> Path:
        if self._registers_dir is not None:
            return self._registers_dir
        return find_existing_registers_dir(self.project_dir)

    @property
    def profit_dir(self) -> Path:
        return find_existing_child_dir(self.project_dir, DEFAULT_PROFIT_DIR)

    @property
    def accounting_report_dir(self) -> Path:
        return find_existing_child_dir(self.project_dir, DEFAULT_ACCOUNTING_REPORT_DIR)

    @property
    def accounting_table(self) -> Path:
        return self.project_dir / "accounting_table.xlsx"

    @property
    def transaction_history(self) -> Path:
        return self.project_dir / "Історія Транзакцій.xlsx"

    @property
    def output_registers_workbook(self) -> Path:
        return self.registers_dir / DEFAULT_OUTPUT_WORKBOOK_NAME

def ensure_project_structure(project_dir: Path) -> ProjectPaths:
    resolved_project_dir = project_dir.expanduser().resolve()
    # Не створюємо project_dir або Реєстри автоматично, щоб не робити паралельні папки.
    if not resolved_project_dir.is_dir():
        raise FileNotFoundError(f"Папка проєкту не існує: {resolved_project_dir}")
    registers_dir = find_existing_registers_dir(resolved_project_dir)
    paths = ProjectPaths(resolved_project_dir, _registers_dir=registers_dir)
    paths.profit_dir.mkdir(parents=True, exist_ok=True)
    paths.accounting_report_dir.mkdir(parents=True, exist_ok=True)
    return paths

def parse_register_pdf_filename(path: Path) -> Optional[dict]:
    match = PDF_REGISTER_RE.match(path.name)
    if not match:
        return None

    kind_raw = match.group("kind").lower()
    kind = "purchase" if "закуп" in kind_raw else "sales"
    start_date = parse_compact_register_date(match.group("start"), is_start=True)
    end_date = parse_compact_register_date(match.group("end"), is_start=False)
    return {
        "File name": path.name,
        "Register type": "Закупівлі" if kind == "purchase" else "Продажу",
        "Period start": start_date,
        "Period end": end_date,
        "Path": str(path),
    }

def scan_register_pdfs(registers_dir: Path) -> pd.DataFrame:
    rows = []
    if registers_dir.exists():
        for pdf in sorted(registers_dir.glob("*.pdf")):
            parsed = parse_register_pdf_filename(pdf)
            if parsed:
                rows.append(parsed)
    return pd.DataFrame(rows, columns=["File name", "Register type", "Period start", "Period end", "Path"])
