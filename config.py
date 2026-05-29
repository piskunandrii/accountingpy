from pathlib import Path

DEFAULT_PROJECT_DIR = Path(__file__).resolve().parent

DEFAULT_REGISTERS_DIR_NAME = "Registers"
DEFAULT_PROFIT_DIR_NAME = "Profit"
DEFAULT_ACCOUNTING_REPORT_DIR_NANE = "Accounting report"

DEFAULT_REGISTERS_DIR = DEFAULT_PROJECT_DIR / DEFAULT_REGISTERS_DIR_NAME
DEFAULT_PROFIT_DIR = DEFAULT_PROJECT_DIR / DEFAULT_PROFIT_DIR_NAME
DEFAULT_ACCOUNTING_REPORT_DIR = DEFAULT_PROJECT_DIR / DEFAULT_ACCOUNTING_REPORT_DIR_NANE

DEFAULT_PROJECT_DIR = "/Users/andriiipiskun/Desktop/ISS-Chem sp. z o. o."
DEFAULT_ACCOUNTING_TABLE = "accounting_table.xlsx"
DEFAULT_REPORTS_DIR = "Бухгалтерський звіт"
DEFAULT_OUTPUT_WORKBOOK_NAME = "реєстри.xlsx"

NBP_API = "https://api.nbp.pl/api/exchangerates/rates/a/{currency}/{date}/?format=json"
CACHE_FILE = DEFAULT_PROJECT_DIR / "nbp_cache.json"
