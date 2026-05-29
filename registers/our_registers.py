import json
import math
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, Iterable, Optional

import pandas as pd
import requests

from config import NBP_API, CACHE_FILE, NORMALIZE_FV_TO_FE, ROUND_PLN_TO
from core.numbers import clean_text, clean_doc_number, normalize_currency, to_num
from core.dates import to_date, previous_calendar_day, date_chunks

import pytesseract
from PIL import Image
from openpyxl.utils import get_column_letter


NBP_API = "https://api.nbp.pl/api/exchangerates/rates/a/eur/{start}/{end}/?format=json"
CACHE_FILE = "nbp_eur_rates_cache.json"

NORMALIZE_FV_TO_FE = True
ROUND_PLN_TO = 2

SALES_OUTPUT_FILE = "rejestr_sprzedazy.xlsx"
PURCHASE_OUTPUT_FILE = "rejestr_zakupu.xlsx"


SALES_REQUIRED_COLUMNS = {
    "BuyFS": [
        "Date of issue",
        "Invoice number",
        "Customer",
        "Invoice Net Amount",
        "Invoice Vat Amount",
        "Invoice Gross Amount",
        "Advance invoice number",
        "Final invoice number",
    ],
    "BuyFZ": [
        "Date of issue",
        "Advance invoice number",
        "Customer",
        "Advance invoice Net Amount",
        "Advance Invoice Vat Amount",
        "Advance invoice Gross Amount",
    ],
    "BuyFK": [
        "Date of issue",
        "Final invoice number",
        "Customer",
        "Net DC amount",
        "Vat DC amount",
        "Gross DC amount",
        "Advance invoice number",
        "Sales invoices numbers",
    ],
}


PURCHASE_REQUIRED_COLUMNS = {
    "SupFS": [
        "Date of issue",
        "Supplier's invoice number",
        "Supplier",
        "Invoice Net Amount",
        "Invoice Vat Amount",
        "Invoice Gross Amount",
    ],
    "ServicesFS": [
        "Date of issue",
        "Service invoice number",
        "Service company",
        "Invoice Net Amount",
        "Invoice Vat Amount",
        "Invoice Gross Amount",
        "Invoice currency",
    ],
}

def load_cache(cache_path: Path) -> Dict[str, float]:
    if not cache_path.exists():
        return {}
    with cache_path.open("r", encoding="utf-8") as f:
        raw = json.load(f)
    return {str(k): float(v) for k, v in raw.items()}

def save_cache(cache_path: Path, rates: Dict[str, float]) -> None:
    with cache_path.open("w", encoding="utf-8") as f:
        json.dump(rates, f, ensure_ascii=False, indent=2, sort_keys=True)

def fetch_nbp_rates(start: date, end: date, cache_path: Path) -> Dict[date, float]:
    """
    Fetch EUR/PLN rates from NBP table A.
    Returns mapping: rate_date -> EUR/PLN mid rate.
    """
    cached = load_cache(cache_path)

    for chunk_start, chunk_end in date_chunks(start, end):
        # We fetch a chunk when at least one calendar date in that chunk is not cached.
        # Calendar dates not returned by NBP are weekends/holidays, so this is intentionally broad.
        missing_any = False
        cursor = chunk_start
        while cursor <= chunk_end:
            if cursor.isoformat() not in cached:
                missing_any = True
                break
            cursor += timedelta(days=1)

        if missing_any:
            url = NBP_API.format(start=chunk_start.isoformat(), end=chunk_end.isoformat())
            response = requests.get(url, timeout=30)
            if response.status_code == 404:
                continue
            response.raise_for_status()
            data = response.json()
            for item in data.get("rates", []):
                cached[item["effectiveDate"]] = float(item["mid"])
            time.sleep(0.15)

    save_cache(cache_path, cached)
    return {datetime.fromisoformat(k).date(): v for k, v in cached.items()}

def latest_rate_before(invoice_date: date, rates: Dict[date, float]) -> tuple[date, float]:
    """
    For invoice date D, return latest available NBP rate date < D.
    This implements previous business day / last available before issue date.
    """
    d = previous_calendar_day(invoice_date)
    for _ in range(31):
        if d in rates:
            return d, rates[d]
        d -= timedelta(days=1)
    raise ValueError(f"No NBP EUR rate found before {invoice_date.isoformat()}")

def get_rates_for_dates(issue_dates: list[date], xlsx_path: Path) -> Dict[date, float]:
    if not issue_dates:
        return {}
    min_rate_date = min(issue_dates) - timedelta(days=31)
    max_rate_date = max(issue_dates) - timedelta(days=1)
    return fetch_nbp_rates(min_rate_date, max_rate_date, xlsx_path.with_name(CACHE_FILE))

def validate_workbook(xlsx_path: Path, required_columns: dict[str, list[str]]) -> None:
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Input file not found: {xlsx_path}")

    xl = pd.ExcelFile(xlsx_path)

    missing_sheets = [s for s in required_columns if s not in xl.sheet_names]
    if missing_sheets:
        raise ValueError(f"Missing sheets: {missing_sheets}")

    for sheet_name, required in required_columns.items():
        cols = pd.read_excel(xlsx_path, sheet_name=sheet_name, nrows=0).columns.tolist()
        missing_cols = [c for c in required if c not in cols]
        if missing_cols:
            raise ValueError(f"Sheet {sheet_name} is missing columns: {missing_cols}")

def add_sales_pln_fields(
    record: dict,
    issue_date: date,
    eur_net: float,
    eur_vat: float,
    eur_gross: float,
    rates: Dict[date, float],
) -> dict:
    rate_date, rate = latest_rate_before(issue_date, rates)
    record.update(
        {
            "NBP rate date": rate_date,
            "NBP EUR/PLN rate": rate,
            "Net EUR source": eur_net,
            "VAT EUR source": eur_vat,
            "Gross EUR source": eur_gross,
            "Netto PLN": round(eur_net * rate, ROUND_PLN_TO),
            "VAT PLN": round(eur_vat * rate, ROUND_PLN_TO),
            "Brutto PLN": round(eur_gross * rate, ROUND_PLN_TO),
        }
    )
    return record

def generate_sales_register(xlsx_path: Path) -> pd.DataFrame:
    validate_workbook(xlsx_path, SALES_REQUIRED_COLUMNS)

    buyfs = pd.read_excel(xlsx_path, sheet_name="BuyFS")
    buyfz = pd.read_excel(xlsx_path, sheet_name="BuyFZ")
    buyfk = pd.read_excel(xlsx_path, sheet_name="BuyFK")

    all_dates: list[date] = []
    for df in [buyfs, buyfz, buyfk]:
        dates = pd.to_datetime(df["Date of issue"], errors="coerce").dropna().dt.date.tolist()
        all_dates.extend(dates)

    if not all_dates:
        raise ValueError("No valid Date of issue values found for sales register.")

    rates = get_rates_for_dates(all_dates, xlsx_path)
    records = []

    # 1) BuyFS -> FE/FV sales invoices
    for _, row in buyfs.iterrows():
        if pd.isna(row.get("Invoice number")):
            continue

        issue_date = to_date(row["Date of issue"])
        doc_number = clean_doc_number(row["Invoice number"])
        linked_advance = clean_doc_number(row.get("Advance invoice number"))
        linked_final = clean_doc_number(row.get("Final invoice number"))
        is_linked_to_fz_fk = bool(linked_advance or linked_final)

        if is_linked_to_fz_fk:
            eur_net = eur_vat = eur_gross = 0.0
            logic = "FE linked to FZ/FK -> 0 PLN to avoid double counting"
        else:
            eur_net = to_num(row.get("Invoice Net Amount"))
            eur_vat = to_num(row.get("Invoice Vat Amount"))
            eur_gross = to_num(row.get("Invoice Gross Amount"))
            logic = "Standalone FE -> invoice amount converted EUR to PLN"

        rec = {
            "Source sheet": "BuyFS",
            "Document type": "FE",
            "Date of issue": issue_date,
            "Nr dokumentu": doc_number,
            "Customer": row.get("Customer", ""),
            "Advance invoice number": linked_advance,
            "Final invoice number": linked_final,
            "Sales invoices numbers": "",
            "Accounting logic": logic,
        }
        records.append(add_sales_pln_fields(rec, issue_date, eur_net, eur_vat, eur_gross, rates))

    # 2) BuyFZ -> advance invoices
    for _, row in buyfz.iterrows():
        if pd.isna(row.get("Advance invoice number")):
            continue

        issue_date = to_date(row["Date of issue"])

        # IMPORTANT: use Advance invoice amounts, not Order net amount.
        eur_net = to_num(row.get("Advance invoice Net Amount"))
        eur_vat = to_num(row.get("Advance Invoice Vat Amount"))
        eur_gross = to_num(row.get("Advance invoice Gross Amount"))

        rec = {
            "Source sheet": "BuyFZ",
            "Document type": "FZ",
            "Date of issue": issue_date,
            "Nr dokumentu": clean_doc_number(row["Advance invoice number"]),
            "Customer": row.get("Customer", ""),
            "Advance invoice number": clean_doc_number(row.get("Advance invoice number")),
            "Final invoice number": clean_doc_number(row.get("Final invoice number")),
            "Sales invoices numbers": "",
            "Accounting logic": "FZ advance invoice -> advance amount converted EUR to PLN",
        }
        records.append(add_sales_pln_fields(rec, issue_date, eur_net, eur_vat, eur_gross, rates))

    # 3) BuyFK -> final invoices / difference corrections
    for _, row in buyfk.iterrows():
        if pd.isna(row.get("Final invoice number")):
            continue

        issue_date = to_date(row["Date of issue"])
        eur_net = to_num(row.get("Net DC amount"))
        eur_vat = to_num(row.get("Vat DC amount"))
        eur_gross = to_num(row.get("Gross DC amount"))

        rec = {
            "Source sheet": "BuyFK",
            "Document type": "FK",
            "Date of issue": issue_date,
            "Nr dokumentu": clean_doc_number(row["Final invoice number"]),
            "Customer": row.get("Customer", ""),
            "Advance invoice number": clean_doc_number(row.get("Advance invoice number")),
            "Final invoice number": clean_doc_number(row.get("Final invoice number")),
            "Sales invoices numbers": row.get("Sales invoices numbers", ""),
            "Accounting logic": "FK final invoice -> only DC difference converted EUR to PLN",
        }
        records.append(add_sales_pln_fields(rec, issue_date, eur_net, eur_vat, eur_gross, rates))

    result = pd.DataFrame(records)
    result = result.sort_values(["Date of issue", "Document type", "Nr dokumentu"]).reset_index(drop=True)
    result.insert(0, "Lp", range(1, len(result) + 1))

    ordered_cols = [
        "Lp",
        "Date of issue",
        "Nr dokumentu",
        "Customer",
        "Netto PLN",
        "VAT PLN",
        "Brutto PLN",
        "Document type",
        "NBP rate date",
        "NBP EUR/PLN rate",
        "Net EUR source",
        "VAT EUR source",
        "Gross EUR source",
        "Source sheet",
        "Advance invoice number",
        "Final invoice number",
        "Sales invoices numbers",
        "Accounting logic",
    ]
    return result[ordered_cols]

def add_purchase_pln_fields(
    record: dict,
    issue_date: date,
    original_currency: str,
    original_net: float,
    original_vat: float,
    original_gross: float,
    rates: Dict[date, float],
) -> dict:
    if original_currency == "PLN":
        rate_date = None
        rate = 1.0
        netto_pln = original_net
        vat_pln = original_vat
        brutto_pln = original_gross
        status = "OK"
        warning = ""
    elif original_currency == "EUR":
        rate_date, rate = latest_rate_before(issue_date, rates)
        netto_pln = original_net * rate
        vat_pln = original_vat * rate
        brutto_pln = original_gross * rate
        status = "OK"
        warning = ""
    else:
        rate_date = None
        rate = None
        netto_pln = None
        vat_pln = None
        brutto_pln = None
        status = "PROBLEM"
        warning = f"Unsupported invoice currency: {original_currency or 'EMPTY'}"

    record.update(
        {
            "Netto PLN": round(netto_pln, ROUND_PLN_TO) if netto_pln is not None else None,
            "VAT PLN": round(vat_pln, ROUND_PLN_TO) if vat_pln is not None else None,
            "Brutto PLN": round(brutto_pln, ROUND_PLN_TO) if brutto_pln is not None else None,
            "Source currency": original_currency,
            "NBP rate date": rate_date,
            "NBP EUR/PLN rate": rate,
            "Original net amount": original_net,
            "Original VAT amount": original_vat,
            "Original gross amount": original_gross,
            "Status": status,
            "Warning": warning,
        }
    )
    return record

def generate_purchase_register(xlsx_path: Path) -> pd.DataFrame:
    validate_workbook(xlsx_path, PURCHASE_REQUIRED_COLUMNS)

    supfs = pd.read_excel(xlsx_path, sheet_name="SupFS")
    services = pd.read_excel(xlsx_path, sheet_name="ServicesFS")

    all_dates: list[date] = []
    for df in [supfs, services]:
        dates = pd.to_datetime(df["Date of issue"], errors="coerce").dropna().dt.date.tolist()
        all_dates.extend(dates)

    if not all_dates:
        raise ValueError("No valid Date of issue values found for purchase register.")

    rates = get_rates_for_dates(all_dates, xlsx_path)
    records = []

    # 1) SupFS - all amounts are treated as EUR according to the requested logic.
    for _, row in supfs.iterrows():
        if pd.isna(row.get("Supplier's invoice number")):
            continue

        issue_date = to_date(row["Date of issue"])
        original_currency = "EUR"

        rec = {
            "Date of issue": issue_date,
            "Invoice number": clean_text(row.get("Supplier's invoice number")),
            "Supplier / Company": row.get("Supplier", ""),
            "Source sheet": "SupFS",
        }
        records.append(
            add_purchase_pln_fields(
                rec,
                issue_date,
                original_currency,
                to_num(row.get("Invoice Net Amount")),
                to_num(row.get("Invoice Vat Amount")),
                to_num(row.get("Invoice Gross Amount")),
                rates,
            )
        )

    # 2) ServicesFS - depends on Invoice currency.
    for _, row in services.iterrows():
        if pd.isna(row.get("Service invoice number")):
            continue

        issue_date = to_date(row["Date of issue"])
        original_currency = normalize_currency(row.get("Invoice currency"))

        rec = {
            "Date of issue": issue_date,
            "Invoice number": clean_text(row.get("Service invoice number")),
            "Supplier / Company": row.get("Service company", ""),
            "Source sheet": "ServicesFS",
        }
        records.append(
            add_purchase_pln_fields(
                rec,
                issue_date,
                original_currency,
                to_num(row.get("Invoice Net Amount")),
                to_num(row.get("Invoice Vat Amount")),
                to_num(row.get("Invoice Gross Amount")),
                rates,
            )
        )

    result = pd.DataFrame(records)
    result = result.sort_values(["Date of issue", "Source sheet", "Invoice number"]).reset_index(drop=True)
    result.insert(0, "Lp", range(1, len(result) + 1))

    ordered_cols = [
        "Lp",
        "Date of issue",
        "Invoice number",
        "Supplier / Company",
        "Netto PLN",
        "VAT PLN",
        "Brutto PLN",
        "Source sheet",
        "Source currency",
        "NBP rate date",
        "NBP EUR/PLN rate",
        "Original net amount",
        "Original VAT amount",
        "Original gross amount",
        "Status",
        "Warning",
    ]
    return result[ordered_cols]

def add_sales_summary_sheet(writer: pd.ExcelWriter, register: pd.DataFrame) -> None:
    summary = (
        register.groupby("Document type", as_index=False)
        .agg(
            Rows=("Nr dokumentu", "count"),
            Netto_PLN=("Netto PLN", "sum"),
            VAT_PLN=("VAT PLN", "sum"),
            Brutto_PLN=("Brutto PLN", "sum"),
        )
        .sort_values("Document type")
    )
    total = pd.DataFrame(
        [{
            "Document type": "TOTAL",
            "Rows": int(summary["Rows"].sum()),
            "Netto_PLN": summary["Netto_PLN"].sum(),
            "VAT_PLN": summary["VAT_PLN"].sum(),
            "Brutto_PLN": summary["Brutto_PLN"].sum(),
        }]
    )
    summary = pd.concat([summary, total], ignore_index=True)
    summary.to_excel(writer, sheet_name="Summary", index=False)

def add_purchase_summary_sheet(writer: pd.ExcelWriter, register: pd.DataFrame) -> None:
    summary = (
        register.groupby(["Source sheet", "Source currency", "Status"], dropna=False, as_index=False)
        .agg(
            Rows=("Invoice number", "count"),
            Netto_PLN=("Netto PLN", "sum"),
            VAT_PLN=("VAT PLN", "sum"),
            Brutto_PLN=("Brutto PLN", "sum"),
        )
        .sort_values(["Source sheet", "Source currency", "Status"])
    )
    total = pd.DataFrame(
        [{
            "Source sheet": "TOTAL",
            "Source currency": "",
            "Status": "",
            "Rows": int(summary["Rows"].sum()),
            "Netto_PLN": summary["Netto_PLN"].sum(),
            "VAT_PLN": summary["VAT_PLN"].sum(),
            "Brutto_PLN": summary["Brutto_PLN"].sum(),
        }]
    )
    summary = pd.concat([summary, total], ignore_index=True)
    summary.to_excel(writer, sheet_name="Summary", index=False)

    problems = register[register["Status"].ne("OK") | register["Warning"].fillna("").ne("")]
    if not problems.empty:
        problems.to_excel(writer, sheet_name="Problems", index=False)

def build_sales_summary_df(register: pd.DataFrame) -> pd.DataFrame:
    summary = (
        register.groupby("Document type", as_index=False)
        .agg(
            Rows=("Nr dokumentu", "count"),
            Netto_PLN=("Netto PLN", "sum"),
            VAT_PLN=("VAT PLN", "sum"),
            Brutto_PLN=("Brutto PLN", "sum"),
        )
        .sort_values("Document type")
    )
    total = pd.DataFrame([{
        "Document type": "TOTAL",
        "Rows": int(summary["Rows"].sum()),
        "Netto_PLN": summary["Netto_PLN"].sum(),
        "VAT_PLN": summary["VAT_PLN"].sum(),
        "Brutto_PLN": summary["Brutto_PLN"].sum(),
    }])
    return pd.concat([summary, total], ignore_index=True)

def build_purchase_summary_df(register: pd.DataFrame) -> pd.DataFrame:
    summary = (
        register.groupby(["Source sheet", "Source currency", "Status"], dropna=False, as_index=False)
        .agg(
            Rows=("Invoice number", "count"),
            Netto_PLN=("Netto PLN", "sum"),
            VAT_PLN=("VAT PLN", "sum"),
            Brutto_PLN=("Brutto PLN", "sum"),
        )
        .sort_values(["Source sheet", "Source currency", "Status"])
    )
    total = pd.DataFrame([{
        "Source sheet": "TOTAL",
        "Source currency": "",
        "Status": "",
        "Rows": int(summary["Rows"].sum()),
        "Netto_PLN": summary["Netto_PLN"].sum(),
        "VAT_PLN": summary["VAT_PLN"].sum(),
        "Brutto_PLN": summary["Brutto_PLN"].sum(),
    }])
    return pd.concat([summary, total], ignore_index=True)
