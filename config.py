from pathlib import Path

DEFAULT_PROJECT_DIR = Path(__file__).resolve().parent.parent

DEFAULT_REGISTERS_DIR_NAME = "Реєстри"
DEFAULT_PROFIT_DIR_NAME = "Прибуток"
DEFAULT_ACCOUNTING_REPORT_DIR_NAME = "Бухгалтерський звіт"

DEFAULT_REGISTERS_DIR = DEFAULT_REGISTERS_DIR_NAME
DEFAULT_PROFIT_DIR = DEFAULT_PROFIT_DIR_NAME
DEFAULT_ACCOUNTING_REPORT_DIR = DEFAULT_ACCOUNTING_REPORT_DIR_NAME
DEFAULT_ACCOUNTING_TABLE = "accounting_table.xlsx"
DEFAULT_REPORTS_DIR = "Бухгалтерський звіт"
DEFAULT_OUTPUT_WORKBOOK_NAME = "реєстри.xlsx"

NBP_API = "https://api.nbp.pl/api/exchangerates/rates/a/{currency}/{date}/?format=json"
CACHE_FILE = DEFAULT_PROJECT_DIR / "nbp_cache.json"
NORMALIZE_FV_TO_FE = True
ROUND_PLN_TO = 2
