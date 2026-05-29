# """
# GUI_VERSION = "2026-05-29-no-source-xlsx-field"
# Generate accounting registers from Iss-chem workbook.
#
# Input workbook:
#     Iss-chem 01.01.2023-12.05.2026.xlsx
#
# Available outputs:
#     1. Rejestr sprzedazy (sales register)
#        Used sheets:
#            BuyFS  - sales invoices in EUR
#            BuyFZ  - advance invoices in EUR
#            BuyFK  - final invoices / difference corrections in EUR
#
#     2. Rejestr zakupu (purchase register)
#        Used sheets:
#            SupFS       - supplier invoices, amounts in EUR
#            ServicesFS  - service invoices, currency depends on "Invoice currency"
#
# NBP rate logic:
#     For invoice date D, use the latest available official NBP EUR/PLN rate
#     strictly before D. This covers weekends and public holidays automatically.
#
# Run:
#     python main.py
#
# Optional:
#     python main.py --register sales
#     python main.py --register purchase
#     python main.py "Iss-chem 01.01.2023-12.05.2026.xlsx"
# """
#
# from __future__ import annotations
#
# import argparse
# import json
# import math
# import re
# import time
# from datetime import date, datetime, timedelta
# from pathlib import Path
# from typing import Dict, Iterable, Optional
# from copy import copy
#
# import pandas as pd
# import requests
#
#
# NBP_API = "https://api.nbp.pl/api/exchangerates/rates/a/eur/{start}/{end}/?format=json"
# CACHE_FILE = "nbp_eur_rates_cache.json"
#
# NORMALIZE_FV_TO_FE = True
# ROUND_PLN_TO = 2
#
# SALES_OUTPUT_FILE = "rejestr_sprzedazy.xlsx"
# PURCHASE_OUTPUT_FILE = "rejestr_zakupu.xlsx"
#
#
# SALES_REQUIRED_COLUMNS = {
#     "BuyFS": [
#         "Date of issue",
#         "Invoice number",
#         "Customer",
#         "Invoice Net Amount",
#         "Invoice Vat Amount",
#         "Invoice Gross Amount",
#         "Advance invoice number",
#         "Final invoice number",
#     ],
#     "BuyFZ": [
#         "Date of issue",
#         "Advance invoice number",
#         "Customer",
#         "Advance invoice Net Amount",
#         "Advance Invoice Vat Amount",
#         "Advance invoice Gross Amount",
#     ],
#     "BuyFK": [
#         "Date of issue",
#         "Final invoice number",
#         "Customer",
#         "Net DC amount",
#         "Vat DC amount",
#         "Gross DC amount",
#         "Advance invoice number",
#         "Sales invoices numbers",
#     ],
# }
#
#
# PURCHASE_REQUIRED_COLUMNS = {
#     "SupFS": [
#         "Date of issue",
#         "Supplier's invoice number",
#         "Supplier",
#         "Invoice Net Amount",
#         "Invoice Vat Amount",
#         "Invoice Gross Amount",
#     ],
#     "ServicesFS": [
#         "Date of issue",
#         "Service invoice number",
#         "Service company",
#         "Invoice Net Amount",
#         "Invoice Vat Amount",
#         "Invoice Gross Amount",
#         "Invoice currency",
#     ],
# }
#
#
# def clean_text(value) -> str:
#     if pd.isna(value):
#         return ""
#     text = str(value).strip()
#     return re.sub(r"\s+", " ", text)
#
#
# def clean_doc_number(value) -> str:
#     """Normalize document number text for stable joins/export."""
#     text = clean_text(value)
#     if NORMALIZE_FV_TO_FE:
#         text = re.sub(r"^FV\b", "FE", text, flags=re.IGNORECASE)
#     return text
#
#
# def normalize_currency(value, default: Optional[str] = None) -> str:
#     if pd.isna(value) or str(value).strip() == "":
#         return default or ""
#     return str(value).strip().upper()
#
#
# def to_date(value) -> date:
#     if pd.isna(value):
#         raise ValueError("Missing Date of issue")
#     if isinstance(value, datetime):
#         return value.date()
#     if isinstance(value, date):
#         return value
#     return pd.to_datetime(value).date()
#
#
# def to_num(value) -> float:
#     """Convert Excel/Polish/European numeric values to float."""
#     if pd.isna(value) or value == "":
#         return 0.0
#     if isinstance(value, (int, float)) and not isinstance(value, bool):
#         if isinstance(value, float) and math.isnan(value):
#             return 0.0
#         return float(value)
#
#     text = str(value).strip().replace("\u00a0", "")
#     text = text.replace(" ", "")
#
#     # Polish/European format: 1.234,56 -> 1234.56
#     if "," in text and "." in text:
#         text = text.replace(".", "").replace(",", ".")
#     elif "," in text:
#         text = text.replace(",", ".")
#
#     return float(text)
#
#
# def previous_calendar_day(d: date) -> date:
#     return d - timedelta(days=1)
#
#
# def date_chunks(start: date, end: date, max_days: int = 360) -> Iterable[tuple[date, date]]:
#     """NBP API has range limits, so fetch in chunks."""
#     current = start
#     while current <= end:
#         chunk_end = min(current + timedelta(days=max_days), end)
#         yield current, chunk_end
#         current = chunk_end + timedelta(days=1)
#
#
# def load_cache(cache_path: Path) -> Dict[str, float]:
#     if not cache_path.exists():
#         return {}
#     with cache_path.open("r", encoding="utf-8") as f:
#         raw = json.load(f)
#     return {str(k): float(v) for k, v in raw.items()}
#
#
# def save_cache(cache_path: Path, rates: Dict[str, float]) -> None:
#     with cache_path.open("w", encoding="utf-8") as f:
#         json.dump(rates, f, ensure_ascii=False, indent=2, sort_keys=True)
#
#
# def fetch_nbp_rates(start: date, end: date, cache_path: Path) -> Dict[date, float]:
#     """
#     Fetch EUR/PLN rates from NBP table A.
#     Returns mapping: rate_date -> EUR/PLN mid rate.
#     """
#     cached = load_cache(cache_path)
#
#     for chunk_start, chunk_end in date_chunks(start, end):
#         # We fetch a chunk when at least one calendar date in that chunk is not cached.
#         # Calendar dates not returned by NBP are weekends/holidays, so this is intentionally broad.
#         missing_any = False
#         cursor = chunk_start
#         while cursor <= chunk_end:
#             if cursor.isoformat() not in cached:
#                 missing_any = True
#                 break
#             cursor += timedelta(days=1)
#
#         if missing_any:
#             url = NBP_API.format(start=chunk_start.isoformat(), end=chunk_end.isoformat())
#             response = requests.get(url, timeout=30)
#             if response.status_code == 404:
#                 continue
#             response.raise_for_status()
#             data = response.json()
#             for item in data.get("rates", []):
#                 cached[item["effectiveDate"]] = float(item["mid"])
#             time.sleep(0.15)
#
#     save_cache(cache_path, cached)
#     return {datetime.fromisoformat(k).date(): v for k, v in cached.items()}
#
#
# def latest_rate_before(invoice_date: date, rates: Dict[date, float]) -> tuple[date, float]:
#     """
#     For invoice date D, return latest available NBP rate date < D.
#     This implements previous business day / last available before issue date.
#     """
#     d = previous_calendar_day(invoice_date)
#     for _ in range(31):
#         if d in rates:
#             return d, rates[d]
#         d -= timedelta(days=1)
#     raise ValueError(f"No NBP EUR rate found before {invoice_date.isoformat()}")
#
#
# def get_rates_for_dates(issue_dates: list[date], xlsx_path: Path) -> Dict[date, float]:
#     if not issue_dates:
#         return {}
#     min_rate_date = min(issue_dates) - timedelta(days=31)
#     max_rate_date = max(issue_dates) - timedelta(days=1)
#     return fetch_nbp_rates(min_rate_date, max_rate_date, xlsx_path.with_name(CACHE_FILE))
#
#
# def validate_workbook(xlsx_path: Path, required_columns: dict[str, list[str]]) -> None:
#     if not xlsx_path.exists():
#         raise FileNotFoundError(f"Input file not found: {xlsx_path}")
#
#     xl = pd.ExcelFile(xlsx_path)
#
#     missing_sheets = [s for s in required_columns if s not in xl.sheet_names]
#     if missing_sheets:
#         raise ValueError(f"Missing sheets: {missing_sheets}")
#
#     for sheet_name, required in required_columns.items():
#         cols = pd.read_excel(xlsx_path, sheet_name=sheet_name, nrows=0).columns.tolist()
#         missing_cols = [c for c in required if c not in cols]
#         if missing_cols:
#             raise ValueError(f"Sheet {sheet_name} is missing columns: {missing_cols}")
#
#
# def add_sales_pln_fields(
#     record: dict,
#     issue_date: date,
#     eur_net: float,
#     eur_vat: float,
#     eur_gross: float,
#     rates: Dict[date, float],
# ) -> dict:
#     rate_date, rate = latest_rate_before(issue_date, rates)
#     record.update(
#         {
#             "NBP rate date": rate_date,
#             "NBP EUR/PLN rate": rate,
#             "Net EUR source": eur_net,
#             "VAT EUR source": eur_vat,
#             "Gross EUR source": eur_gross,
#             "Netto PLN": round(eur_net * rate, ROUND_PLN_TO),
#             "VAT PLN": round(eur_vat * rate, ROUND_PLN_TO),
#             "Brutto PLN": round(eur_gross * rate, ROUND_PLN_TO),
#         }
#     )
#     return record
#
#
# def generate_sales_register(xlsx_path: Path) -> pd.DataFrame:
#     validate_workbook(xlsx_path, SALES_REQUIRED_COLUMNS)
#
#     buyfs = pd.read_excel(xlsx_path, sheet_name="BuyFS")
#     buyfz = pd.read_excel(xlsx_path, sheet_name="BuyFZ")
#     buyfk = pd.read_excel(xlsx_path, sheet_name="BuyFK")
#
#     all_dates: list[date] = []
#     for df in [buyfs, buyfz, buyfk]:
#         dates = pd.to_datetime(df["Date of issue"], errors="coerce").dropna().dt.date.tolist()
#         all_dates.extend(dates)
#
#     if not all_dates:
#         raise ValueError("No valid Date of issue values found for sales register.")
#
#     rates = get_rates_for_dates(all_dates, xlsx_path)
#     records = []
#
#     # 1) BuyFS -> FE/FV sales invoices
#     for _, row in buyfs.iterrows():
#         if pd.isna(row.get("Invoice number")):
#             continue
#
#         issue_date = to_date(row["Date of issue"])
#         doc_number = clean_doc_number(row["Invoice number"])
#         linked_advance = clean_doc_number(row.get("Advance invoice number"))
#         linked_final = clean_doc_number(row.get("Final invoice number"))
#         is_linked_to_fz_fk = bool(linked_advance or linked_final)
#
#         if is_linked_to_fz_fk:
#             eur_net = eur_vat = eur_gross = 0.0
#             logic = "FE linked to FZ/FK -> 0 PLN to avoid double counting"
#         else:
#             eur_net = to_num(row.get("Invoice Net Amount"))
#             eur_vat = to_num(row.get("Invoice Vat Amount"))
#             eur_gross = to_num(row.get("Invoice Gross Amount"))
#             logic = "Standalone FE -> invoice amount converted EUR to PLN"
#
#         rec = {
#             "Source sheet": "BuyFS",
#             "Document type": "FE",
#             "Date of issue": issue_date,
#             "Nr dokumentu": doc_number,
#             "Customer": row.get("Customer", ""),
#             "Advance invoice number": linked_advance,
#             "Final invoice number": linked_final,
#             "Sales invoices numbers": "",
#             "Accounting logic": logic,
#         }
#         records.append(add_sales_pln_fields(rec, issue_date, eur_net, eur_vat, eur_gross, rates))
#
#     # 2) BuyFZ -> advance invoices
#     for _, row in buyfz.iterrows():
#         if pd.isna(row.get("Advance invoice number")):
#             continue
#
#         issue_date = to_date(row["Date of issue"])
#
#         # IMPORTANT: use Advance invoice amounts, not Order net amount.
#         eur_net = to_num(row.get("Advance invoice Net Amount"))
#         eur_vat = to_num(row.get("Advance Invoice Vat Amount"))
#         eur_gross = to_num(row.get("Advance invoice Gross Amount"))
#
#         rec = {
#             "Source sheet": "BuyFZ",
#             "Document type": "FZ",
#             "Date of issue": issue_date,
#             "Nr dokumentu": clean_doc_number(row["Advance invoice number"]),
#             "Customer": row.get("Customer", ""),
#             "Advance invoice number": clean_doc_number(row.get("Advance invoice number")),
#             "Final invoice number": clean_doc_number(row.get("Final invoice number")),
#             "Sales invoices numbers": "",
#             "Accounting logic": "FZ advance invoice -> advance amount converted EUR to PLN",
#         }
#         records.append(add_sales_pln_fields(rec, issue_date, eur_net, eur_vat, eur_gross, rates))
#
#     # 3) BuyFK -> final invoices / difference corrections
#     for _, row in buyfk.iterrows():
#         if pd.isna(row.get("Final invoice number")):
#             continue
#
#         issue_date = to_date(row["Date of issue"])
#         eur_net = to_num(row.get("Net DC amount"))
#         eur_vat = to_num(row.get("Vat DC amount"))
#         eur_gross = to_num(row.get("Gross DC amount"))
#
#         rec = {
#             "Source sheet": "BuyFK",
#             "Document type": "FK",
#             "Date of issue": issue_date,
#             "Nr dokumentu": clean_doc_number(row["Final invoice number"]),
#             "Customer": row.get("Customer", ""),
#             "Advance invoice number": clean_doc_number(row.get("Advance invoice number")),
#             "Final invoice number": clean_doc_number(row.get("Final invoice number")),
#             "Sales invoices numbers": row.get("Sales invoices numbers", ""),
#             "Accounting logic": "FK final invoice -> only DC difference converted EUR to PLN",
#         }
#         records.append(add_sales_pln_fields(rec, issue_date, eur_net, eur_vat, eur_gross, rates))
#
#     result = pd.DataFrame(records)
#     result = result.sort_values(["Date of issue", "Document type", "Nr dokumentu"]).reset_index(drop=True)
#     result.insert(0, "Lp", range(1, len(result) + 1))
#
#     ordered_cols = [
#         "Lp",
#         "Date of issue",
#         "Nr dokumentu",
#         "Customer",
#         "Netto PLN",
#         "VAT PLN",
#         "Brutto PLN",
#         "Document type",
#         "NBP rate date",
#         "NBP EUR/PLN rate",
#         "Net EUR source",
#         "VAT EUR source",
#         "Gross EUR source",
#         "Source sheet",
#         "Advance invoice number",
#         "Final invoice number",
#         "Sales invoices numbers",
#         "Accounting logic",
#     ]
#     return result[ordered_cols]
#
#
# def add_purchase_pln_fields(
#     record: dict,
#     issue_date: date,
#     original_currency: str,
#     original_net: float,
#     original_vat: float,
#     original_gross: float,
#     rates: Dict[date, float],
# ) -> dict:
#     if original_currency == "PLN":
#         rate_date = None
#         rate = 1.0
#         netto_pln = original_net
#         vat_pln = original_vat
#         brutto_pln = original_gross
#         status = "OK"
#         warning = ""
#     elif original_currency == "EUR":
#         rate_date, rate = latest_rate_before(issue_date, rates)
#         netto_pln = original_net * rate
#         vat_pln = original_vat * rate
#         brutto_pln = original_gross * rate
#         status = "OK"
#         warning = ""
#     else:
#         rate_date = None
#         rate = None
#         netto_pln = None
#         vat_pln = None
#         brutto_pln = None
#         status = "PROBLEM"
#         warning = f"Unsupported invoice currency: {original_currency or 'EMPTY'}"
#
#     record.update(
#         {
#             "Netto PLN": round(netto_pln, ROUND_PLN_TO) if netto_pln is not None else None,
#             "VAT PLN": round(vat_pln, ROUND_PLN_TO) if vat_pln is not None else None,
#             "Brutto PLN": round(brutto_pln, ROUND_PLN_TO) if brutto_pln is not None else None,
#             "Source currency": original_currency,
#             "NBP rate date": rate_date,
#             "NBP EUR/PLN rate": rate,
#             "Original net amount": original_net,
#             "Original VAT amount": original_vat,
#             "Original gross amount": original_gross,
#             "Status": status,
#             "Warning": warning,
#         }
#     )
#     return record
#
#
# def generate_purchase_register(xlsx_path: Path) -> pd.DataFrame:
#     validate_workbook(xlsx_path, PURCHASE_REQUIRED_COLUMNS)
#
#     supfs = pd.read_excel(xlsx_path, sheet_name="SupFS")
#     services = pd.read_excel(xlsx_path, sheet_name="ServicesFS")
#
#     all_dates: list[date] = []
#     for df in [supfs, services]:
#         dates = pd.to_datetime(df["Date of issue"], errors="coerce").dropna().dt.date.tolist()
#         all_dates.extend(dates)
#
#     if not all_dates:
#         raise ValueError("No valid Date of issue values found for purchase register.")
#
#     rates = get_rates_for_dates(all_dates, xlsx_path)
#     records = []
#
#     # 1) SupFS - all amounts are treated as EUR according to the requested logic.
#     for _, row in supfs.iterrows():
#         if pd.isna(row.get("Supplier's invoice number")):
#             continue
#
#         issue_date = to_date(row["Date of issue"])
#         original_currency = "EUR"
#
#         rec = {
#             "Date of issue": issue_date,
#             "Invoice number": clean_text(row.get("Supplier's invoice number")),
#             "Supplier / Company": row.get("Supplier", ""),
#             "Source sheet": "SupFS",
#         }
#         records.append(
#             add_purchase_pln_fields(
#                 rec,
#                 issue_date,
#                 original_currency,
#                 to_num(row.get("Invoice Net Amount")),
#                 to_num(row.get("Invoice Vat Amount")),
#                 to_num(row.get("Invoice Gross Amount")),
#                 rates,
#             )
#         )
#
#     # 2) ServicesFS - depends on Invoice currency.
#     for _, row in services.iterrows():
#         if pd.isna(row.get("Service invoice number")):
#             continue
#
#         issue_date = to_date(row["Date of issue"])
#         original_currency = normalize_currency(row.get("Invoice currency"))
#
#         rec = {
#             "Date of issue": issue_date,
#             "Invoice number": clean_text(row.get("Service invoice number")),
#             "Supplier / Company": row.get("Service company", ""),
#             "Source sheet": "ServicesFS",
#         }
#         records.append(
#             add_purchase_pln_fields(
#                 rec,
#                 issue_date,
#                 original_currency,
#                 to_num(row.get("Invoice Net Amount")),
#                 to_num(row.get("Invoice Vat Amount")),
#                 to_num(row.get("Invoice Gross Amount")),
#                 rates,
#             )
#         )
#
#     result = pd.DataFrame(records)
#     result = result.sort_values(["Date of issue", "Source sheet", "Invoice number"]).reset_index(drop=True)
#     result.insert(0, "Lp", range(1, len(result) + 1))
#
#     ordered_cols = [
#         "Lp",
#         "Date of issue",
#         "Invoice number",
#         "Supplier / Company",
#         "Netto PLN",
#         "VAT PLN",
#         "Brutto PLN",
#         "Source sheet",
#         "Source currency",
#         "NBP rate date",
#         "NBP EUR/PLN rate",
#         "Original net amount",
#         "Original VAT amount",
#         "Original gross amount",
#         "Status",
#         "Warning",
#     ]
#     return result[ordered_cols]
#
#
# def add_sales_summary_sheet(writer: pd.ExcelWriter, register: pd.DataFrame) -> None:
#     summary = (
#         register.groupby("Document type", as_index=False)
#         .agg(
#             Rows=("Nr dokumentu", "count"),
#             Netto_PLN=("Netto PLN", "sum"),
#             VAT_PLN=("VAT PLN", "sum"),
#             Brutto_PLN=("Brutto PLN", "sum"),
#         )
#         .sort_values("Document type")
#     )
#     total = pd.DataFrame(
#         [{
#             "Document type": "TOTAL",
#             "Rows": int(summary["Rows"].sum()),
#             "Netto_PLN": summary["Netto_PLN"].sum(),
#             "VAT_PLN": summary["VAT_PLN"].sum(),
#             "Brutto_PLN": summary["Brutto_PLN"].sum(),
#         }]
#     )
#     summary = pd.concat([summary, total], ignore_index=True)
#     summary.to_excel(writer, sheet_name="Summary", index=False)
#
#
# def add_purchase_summary_sheet(writer: pd.ExcelWriter, register: pd.DataFrame) -> None:
#     summary = (
#         register.groupby(["Source sheet", "Source currency", "Status"], dropna=False, as_index=False)
#         .agg(
#             Rows=("Invoice number", "count"),
#             Netto_PLN=("Netto PLN", "sum"),
#             VAT_PLN=("VAT PLN", "sum"),
#             Brutto_PLN=("Brutto PLN", "sum"),
#         )
#         .sort_values(["Source sheet", "Source currency", "Status"])
#     )
#     total = pd.DataFrame(
#         [{
#             "Source sheet": "TOTAL",
#             "Source currency": "",
#             "Status": "",
#             "Rows": int(summary["Rows"].sum()),
#             "Netto_PLN": summary["Netto_PLN"].sum(),
#             "VAT_PLN": summary["VAT_PLN"].sum(),
#             "Brutto_PLN": summary["Brutto_PLN"].sum(),
#         }]
#     )
#     summary = pd.concat([summary, total], ignore_index=True)
#     summary.to_excel(writer, sheet_name="Summary", index=False)
#
#     problems = register[register["Status"].ne("OK") | register["Warning"].fillna("").ne("")]
#     if not problems.empty:
#         problems.to_excel(writer, sheet_name="Problems", index=False)
#
#
# def autosize_columns(worksheet, dataframe: pd.DataFrame, max_width: int = 42) -> None:
#     """
#     Set Excel column widths safely.
#
#     Important:
#     Do not use pandas Series.map(len) here because some pandas/Python versions
#     can keep numeric objects as float and len(float) raises TypeError.
#     """
#     for idx, col in enumerate(dataframe.columns, start=1):
#         max_len = len(str(col))
#
#         for value in dataframe[col].tolist():
#             if pd.isna(value):
#                 text = ""
#             else:
#                 text = str(value)
#             max_len = max(max_len, len(text))
#
#         width = min(max_len + 2, max_width)
#         worksheet.column_dimensions[worksheet.cell(row=1, column=idx).column_letter].width = width
#
#
# def write_excel_openpyxl(
#     register: pd.DataFrame,
#     output_path: Path,
#     register_sheet_name: str,
#     summary_func,
# ) -> None:
#     """
#     Write output using openpyxl engine.
#     This avoids the extra xlsxwriter dependency.
#     """
#     with pd.ExcelWriter(output_path, engine="openpyxl", date_format="yyyy-mm-dd", datetime_format="yyyy-mm-dd") as writer:
#         register.to_excel(writer, sheet_name=register_sheet_name, index=False)
#         summary_func(writer, register)
#
#         workbook = writer.book
#
#         for sheet_name in workbook.sheetnames:
#             ws = workbook[sheet_name]
#             ws.freeze_panes = "A2"
#             ws.auto_filter.ref = ws.dimensions
#
#             for cell in ws[1]:
#                 font = copy(cell.font)
#                 font.bold = True
#                 cell.font = font
#
#                 fill = copy(cell.fill)
#                 fill.fill_type = "solid"
#                 fill.fgColor = "D9EAF7"
#                 cell.fill = fill
#
#             # Basic number formatting by header names.
#             headers = {cell.value: cell.column for cell in ws[1]}
#             for header, col_idx in headers.items():
#                 col_letter = ws.cell(row=1, column=col_idx).column_letter
#
#                 if header and "Date" in str(header) or header in ["Data", "NBP rate date", "Date of issue"]:
#                     for row in range(2, ws.max_row + 1):
#                         ws[f"{col_letter}{row}"].number_format = "yyyy-mm-dd"
#
#                 if header in [
#                     "Netto PLN", "VAT PLN", "Brutto PLN",
#                     "Net EUR source", "VAT EUR source", "Gross EUR source",
#                     "Original net amount", "Original VAT amount", "Original gross amount",
#                     "Netto_PLN", "VAT_PLN", "Brutto_PLN",
#                 ]:
#                     for row in range(2, ws.max_row + 1):
#                         ws[f"{col_letter}{row}"].number_format = '#,##0.00'
#
#                 if header == "NBP EUR/PLN rate":
#                     for row in range(2, ws.max_row + 1):
#                         ws[f"{col_letter}{row}"].number_format = '0.0000'
#
#             # Autosize.
#             if sheet_name == register_sheet_name:
#                 autosize_columns(ws, register)
#             elif sheet_name == "Summary":
#                 summary_df = pd.read_excel(output_path, sheet_name="Summary") if False else None
#                 for col_idx in range(1, ws.max_column + 1):
#                     ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 18
#             elif sheet_name == "Problems":
#                 autosize_columns(ws, register, max_width=42)
#
#
#
# # =========================
# # GUI HELPERS FOR ISS-CHEM
# # =========================
#
# import threading
# import queue
# import unicodedata
# import subprocess
# import sys
# from dataclasses import dataclass
# from tkinter import Tk, StringVar, BooleanVar, BOTH, END, LEFT, RIGHT, X, Y, filedialog, messagebox
# from tkinter import ttk
#
# try:
#     import openpyxl
# except Exception:  # pragma: no cover
#     openpyxl = None
#
# DEFAULT_PROJECT_DIR = Path("/Users/andriiipiskun/Desktop/ISS-Chem sp. z o. o.")
# DEFAULT_REGISTERS_DIR_NAME = "Реєстри"
# DEFAULT_PROFIT_DIR_NAME = "Прибуток"
# DEFAULT_ACCOUNTING_REPORT_DIR_NAME = "Бухгалтерський звіт"
# DEFAULT_OUTPUT_WORKBOOK_NAME = "реєстри.xlsx"
#
# PDF_REGISTER_RE = re.compile(
#     r"^Реєстр\s+(?P<kind>Закупівлі|Закупівель|Продажу|Продажів)\s+"
#     r"(?P<start>\d{6}|\d{8})-(?P<end>\d{6}|\d{8})\.pdf$",
#     flags=re.IGNORECASE,
# )
#
#
# def _norm_folder_name(value: str) -> str:
#     """Normalize folder names so macOS Unicode variants compare correctly."""
#     return unicodedata.normalize("NFC", value).casefold().strip()
#
#
# def find_existing_registers_dir(project_dir: Path) -> Path:
#     """
#     Return ONLY an already-existing 'Реєстри' folder.
#
#     This function intentionally does NOT create the folder. If the app cannot
#     find the existing registers folder, it raises an error with the exact path
#     it checked. This prevents accidental parallel folders such as Desktop/Реєстри
#     or Реєстри/Реєстри.
#     """
#     project_dir = project_dir.expanduser().resolve()
#     wanted = _norm_folder_name(DEFAULT_REGISTERS_DIR_NAME)
#
#     # If the selected folder is already 'Реєстри', use it directly.
#     if _norm_folder_name(project_dir.name) == wanted:
#         if project_dir.is_dir():
#             return project_dir
#         raise FileNotFoundError(f"Папка Реєстри не існує: {project_dir}")
#
#     # Prefer an existing direct child whose normalized name is 'Реєстри'.
#     if project_dir.exists():
#         for child in project_dir.iterdir():
#             if child.is_dir() and _norm_folder_name(child.name) == wanted:
#                 return child
#
#     expected = project_dir / DEFAULT_REGISTERS_DIR_NAME
#     raise FileNotFoundError(
#         "Не знайшов існуючу папку 'Реєстри'.\n"
#         f"Перевірений шлях: {expected}\n"
#         "У полі 'Папка проєкту' має бути папка ISS-Chem sp. z o. o., "
#         "а не Desktop, AccountingPro або інша папка."
#     )
#
#
# @dataclass
# class ProjectPaths:
#     project_dir: Path
#     _registers_dir: Optional[Path] = None
#
#     @property
#     def registers_dir(self) -> Path:
#         if self._registers_dir is not None:
#             return self._registers_dir
#         return find_existing_registers_dir(self.project_dir)
#
#     @property
#     def profit_dir(self) -> Path:
#         return self.project_dir / DEFAULT_PROFIT_DIR_NAME
#
#     @property
#     def accounting_report_dir(self) -> Path:
#         return self.project_dir / DEFAULT_ACCOUNTING_REPORT_DIR_NAME
#
#     @property
#     def accounting_table(self) -> Path:
#         return self.project_dir / "accounting_table.xlsx"
#
#     @property
#     def transaction_history(self) -> Path:
#         return self.project_dir / "Історія Транзакцій.xlsx"
#
#     @property
#     def output_registers_workbook(self) -> Path:
#         return self.registers_dir / DEFAULT_OUTPUT_WORKBOOK_NAME
#
#
# def ensure_project_structure(project_dir: Path) -> ProjectPaths:
#     resolved_project_dir = project_dir.expanduser().resolve()
#     # Не створюємо project_dir або Реєстри автоматично, щоб не робити паралельні папки.
#     if not resolved_project_dir.is_dir():
#         raise FileNotFoundError(f"Папка проєкту не існує: {resolved_project_dir}")
#     registers_dir = find_existing_registers_dir(resolved_project_dir)
#     paths = ProjectPaths(resolved_project_dir, _registers_dir=registers_dir)
#     paths.profit_dir.mkdir(parents=True, exist_ok=True)
#     paths.accounting_report_dir.mkdir(parents=True, exist_ok=True)
#     return paths
#
#
# def parse_compact_register_date(value: str, *, is_start: bool) -> date:
#     """
#     Supports:
#       MMYYYY   -> first/last day of the month depending on is_start
#       DDMMYYYY -> exact day
#     """
#     value = value.strip()
#     if len(value) == 6:
#         month = int(value[:2])
#         year = int(value[2:])
#         if is_start:
#             return date(year, month, 1)
#         if month == 12:
#             return date(year, 12, 31)
#         return date(year, month + 1, 1) - timedelta(days=1)
#     if len(value) == 8:
#         day = int(value[:2])
#         month = int(value[2:4])
#         year = int(value[4:])
#         return date(year, month, day)
#     raise ValueError(f"Unsupported compact date: {value}")
#
#
# def parse_register_pdf_filename(path: Path) -> Optional[dict]:
#     match = PDF_REGISTER_RE.match(path.name)
#     if not match:
#         return None
#
#     kind_raw = match.group("kind").lower()
#     kind = "purchase" if "закуп" in kind_raw else "sales"
#     start_date = parse_compact_register_date(match.group("start"), is_start=True)
#     end_date = parse_compact_register_date(match.group("end"), is_start=False)
#     return {
#         "File name": path.name,
#         "Register type": "Закупівлі" if kind == "purchase" else "Продажу",
#         "Period start": start_date,
#         "Period end": end_date,
#         "Path": str(path),
#     }
#
#
# def scan_register_pdfs(registers_dir: Path) -> pd.DataFrame:
#     rows = []
#     if registers_dir.exists():
#         for pdf in sorted(registers_dir.glob("*.pdf")):
#             parsed = parse_register_pdf_filename(pdf)
#             if parsed:
#                 rows.append(parsed)
#     return pd.DataFrame(rows, columns=["File name", "Register type", "Period start", "Period end", "Path"])
#
#
# def write_multi_sheet_excel_openpyxl(output_path: Path, sheets: dict[str, pd.DataFrame]) -> None:
#     """Write several DataFrames to one workbook with consistent simple formatting."""
#     output_path.parent.mkdir(parents=True, exist_ok=True)
#     with pd.ExcelWriter(output_path, engine="openpyxl", date_format="yyyy-mm-dd", datetime_format="yyyy-mm-dd") as writer:
#         for sheet_name, df in sheets.items():
#             safe_name = sheet_name[:31]
#             df.to_excel(writer, sheet_name=safe_name, index=False)
#
#         workbook = writer.book
#         for sheet_name in workbook.sheetnames:
#             ws = workbook[sheet_name]
#             ws.freeze_panes = "A2"
#             ws.auto_filter.ref = ws.dimensions
#
#             for cell in ws[1]:
#                 font = copy(cell.font)
#                 font.bold = True
#                 cell.font = font
#                 fill = copy(cell.fill)
#                 fill.fill_type = "solid"
#                 fill.fgColor = "D9EAF7"
#                 cell.fill = fill
#
#             headers = {cell.value: cell.column for cell in ws[1]}
#             for header, col_idx in headers.items():
#                 col_letter = ws.cell(row=1, column=col_idx).column_letter
#                 header_text = str(header or "")
#                 if "Date" in header_text or header in ["Data", "NBP rate date", "Date of issue", "Period start", "Period end"]:
#                     for row in range(2, ws.max_row + 1):
#                         ws[f"{col_letter}{row}"].number_format = "yyyy-mm-dd"
#                 if header in [
#                     "Netto PLN", "VAT PLN", "Brutto PLN",
#                     "Net EUR source", "VAT EUR source", "Gross EUR source",
#                     "Original net amount", "Original VAT amount", "Original gross amount",
#                     "Netto_PLN", "VAT_PLN", "Brutto_PLN",
#                 ]:
#                     for row in range(2, ws.max_row + 1):
#                         ws[f"{col_letter}{row}"].number_format = "#,##0.00"
#                 if header == "NBP EUR/PLN rate":
#                     for row in range(2, ws.max_row + 1):
#                         ws[f"{col_letter}{row}"].number_format = "0.0000"
#
#             for col_idx in range(1, ws.max_column + 1):
#                 max_len = 10
#                 for row_idx in range(1, min(ws.max_row, 200) + 1):
#                     value = ws.cell(row=row_idx, column=col_idx).value
#                     max_len = max(max_len, len(str(value)) if value is not None else 0)
#                 ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 46)
#
#
# def build_sales_summary_df(register: pd.DataFrame) -> pd.DataFrame:
#     summary = (
#         register.groupby("Document type", as_index=False)
#         .agg(
#             Rows=("Nr dokumentu", "count"),
#             Netto_PLN=("Netto PLN", "sum"),
#             VAT_PLN=("VAT PLN", "sum"),
#             Brutto_PLN=("Brutto PLN", "sum"),
#         )
#         .sort_values("Document type")
#     )
#     total = pd.DataFrame([{
#         "Document type": "TOTAL",
#         "Rows": int(summary["Rows"].sum()),
#         "Netto_PLN": summary["Netto_PLN"].sum(),
#         "VAT_PLN": summary["VAT_PLN"].sum(),
#         "Brutto_PLN": summary["Brutto_PLN"].sum(),
#     }])
#     return pd.concat([summary, total], ignore_index=True)
#
#
# def build_purchase_summary_df(register: pd.DataFrame) -> pd.DataFrame:
#     summary = (
#         register.groupby(["Source sheet", "Source currency", "Status"], dropna=False, as_index=False)
#         .agg(
#             Rows=("Invoice number", "count"),
#             Netto_PLN=("Netto PLN", "sum"),
#             VAT_PLN=("VAT PLN", "sum"),
#             Brutto_PLN=("Brutto PLN", "sum"),
#         )
#         .sort_values(["Source sheet", "Source currency", "Status"])
#     )
#     total = pd.DataFrame([{
#         "Source sheet": "TOTAL",
#         "Source currency": "",
#         "Status": "",
#         "Rows": int(summary["Rows"].sum()),
#         "Netto_PLN": summary["Netto_PLN"].sum(),
#         "VAT_PLN": summary["VAT_PLN"].sum(),
#         "Brutto_PLN": summary["Brutto_PLN"].sum(),
#     }])
#     return pd.concat([summary, total], ignore_index=True)
#
#
# class RejestrApp:
#     def __init__(self, root: Tk):
#         self.root = root
#         self.root.title("ISS-Chem бухгалтерський helper — Реєстри")
#         self.root.geometry("980x680")
#
#         self.project_dir = StringVar(value=str(DEFAULT_PROJECT_DIR))
#         self.output_xlsx = StringVar(value="")
#         self.generate_sales_var = BooleanVar(value=True)
#         self.generate_purchase_var = BooleanVar(value=True)
#         self.scan_pdf_var = BooleanVar(value=True)
#         self.log_queue: queue.Queue[str] = queue.Queue()
#
#         self._build_ui()
#         self._on_project_changed()
#         self._poll_log_queue()
#
#     def _build_ui(self) -> None:
#         pad = {"padx": 10, "pady": 6}
#         main = ttk.Frame(self.root)
#         main.pack(fill=BOTH, expand=True)
#
#         paths_box = ttk.LabelFrame(main, text="Шляхи")
#         paths_box.pack(fill=X, **pad)
#
#         self._path_row(paths_box, "Папка проєкту", self.project_dir, self.choose_project_dir, 0)
#         self._path_row(paths_box, "Файл результату", self.output_xlsx, self.choose_output_xlsx, 1)
#
#         options_box = ttk.LabelFrame(main, text="Що зробити")
#         options_box.pack(fill=X, **pad)
#         ttk.Checkbutton(options_box, text="Згенерувати Реєстр продажу", variable=self.generate_sales_var).pack(side=LEFT, padx=10, pady=8)
#         ttk.Checkbutton(options_box, text="Згенерувати Реєстр закупівлі", variable=self.generate_purchase_var).pack(side=LEFT, padx=10, pady=8)
#         ttk.Checkbutton(options_box, text="Додати список PDF-реєстрів з папки ‘Реєстри’", variable=self.scan_pdf_var).pack(side=LEFT, padx=10, pady=8)
#
#         actions = ttk.Frame(main)
#         actions.pack(fill=X, **pad)
#         ttk.Button(actions, text="1. Перевірити структуру папок", command=self.create_structure).pack(side=LEFT, padx=5)
#         ttk.Button(actions, text="2. Запустити генерацію", command=self.run_generation_threaded).pack(side=LEFT, padx=5)
#         ttk.Button(actions, text="Відкрити папку Реєстри", command=self.open_registers_dir).pack(side=LEFT, padx=5)
#         ttk.Button(actions, text="Очистити лог", command=lambda: self.log_text.delete("1.0", END)).pack(side=RIGHT, padx=5)
#
#         log_box = ttk.LabelFrame(main, text="Лог")
#         log_box.pack(fill=BOTH, expand=True, **pad)
#         self.log_text = __import__("tkinter").Text(log_box, wrap="word", height=22)
#         scrollbar = ttk.Scrollbar(log_box, command=self.log_text.yview)
#         self.log_text.configure(yscrollcommand=scrollbar.set)
#         self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
#         scrollbar.pack(side=RIGHT, fill=Y)
#
#         hint = (
#             "Назви PDF підтримуються так: ‘Реєстр Закупівлі 012026-032026.pdf’ "
#             "або ‘Реєстр Закупівлі 01012026-15032026.pdf’. "
#             "MMYYYY означає весь місяць, DDMMYYYY — точну дату."
#         )
#         ttk.Label(main, text=hint, foreground="#555").pack(fill=X, padx=14, pady=(0, 8))
#
#     def _path_row(self, parent, label: str, var: StringVar, command, row: int) -> None:
#         ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=5)
#         entry = ttk.Entry(parent, textvariable=var)
#         entry.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
#         ttk.Button(parent, text="Обрати…", command=command).grid(row=row, column=2, padx=8, pady=5)
#         parent.columnconfigure(1, weight=1)
#
#     def log(self, message: str) -> None:
#         timestamp = datetime.now().strftime("%H:%M:%S")
#         self.log_queue.put(f"[{timestamp}] {message}\n")
#
#     def _poll_log_queue(self) -> None:
#         try:
#             while True:
#                 msg = self.log_queue.get_nowait()
#                 self.log_text.insert(END, msg)
#                 self.log_text.see(END)
#         except queue.Empty:
#             pass
#         self.root.after(150, self._poll_log_queue)
#
#     def _on_project_changed(self) -> None:
#         project = Path(self.project_dir.get()).expanduser()
#         paths = ProjectPaths(project)
#         self.output_xlsx.set(str(paths.output_registers_workbook))
#     def choose_project_dir(self) -> None:
#         selected = filedialog.askdirectory(initialdir=str(Path.home() / "Desktop"), title="Оберіть папку ISS-Chem sp. z o. o.")
#         if selected:
#             selected_path = Path(selected).expanduser()
#             # If the user accidentally chooses the Реєстри folder itself, use its parent
#             # as the project folder. Otherwise the app would create Реєстри/Реєстри.
#             if _norm_folder_name(selected_path.name) == _norm_folder_name(DEFAULT_REGISTERS_DIR_NAME):
#                 selected_path = selected_path.parent
#             self.project_dir.set(str(selected_path))
#             self._on_project_changed()
#
#     def choose_output_xlsx(self) -> None:
#         selected = filedialog.asksaveasfilename(
#             initialdir=str(ProjectPaths(Path(self.project_dir.get())).registers_dir),
#             initialfile=DEFAULT_OUTPUT_WORKBOOK_NAME,
#             defaultextension=".xlsx",
#             title="Куди зберегти файл реєстрів",
#             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
#         )
#         if selected:
#             self.output_xlsx.set(selected)
#
#     def create_structure(self) -> None:
#         try:
#             paths = ensure_project_structure(Path(self.project_dir.get()))
#             self.output_xlsx.set(str(paths.output_registers_workbook))
#             self.log(f"Структуру папок перевірено: {paths.project_dir}")
#             self.log(f"Реєстри: {paths.registers_dir}")
#             self.log(f"Прибуток: {paths.profit_dir}")
#             self.log(f"Бухгалтерський звіт: {paths.accounting_report_dir}")
#         except Exception as exc:
#             messagebox.showerror("Помилка", str(exc))
#             self.log(f"ПОМИЛКА: {exc}")
#
#     def run_generation_threaded(self) -> None:
#         thread = threading.Thread(target=self.run_generation, daemon=True)
#         thread.start()
#
#     def run_generation(self) -> None:
#         try:
#             paths = ensure_project_structure(Path(self.project_dir.get()))
#             source = paths.accounting_table.expanduser().resolve()
#             output = Path(self.output_xlsx.get()).expanduser().resolve()
#
#             if not source.exists():
#                 raise FileNotFoundError(f"Не знайдено accounting_table.xlsx: {source}")
#             if not (self.generate_sales_var.get() or self.generate_purchase_var.get() or self.scan_pdf_var.get()):
#                 raise ValueError("Оберіть хоча б одну дію.")
#
#             self.log(f"Старт. Дані беру автоматично з: {source}")
#             sheets: dict[str, pd.DataFrame] = {}
#
#             if self.generate_sales_var.get():
#                 self.log("Генерую Реєстр продажу…")
#                 sales = generate_sales_register(source)
#                 sheets["Rejestr sprzedazy"] = sales
#                 sheets["Summary sprzedazy"] = build_sales_summary_df(sales)
#                 self.log(f"Реєстр продажу готовий: {len(sales)} рядків.")
#
#             if self.generate_purchase_var.get():
#                 self.log("Генерую Реєстр закупівлі…")
#                 purchase = generate_purchase_register(source)
#                 sheets["Rejestr zakupu"] = purchase
#                 sheets["Summary zakupu"] = build_purchase_summary_df(purchase)
#                 problems = purchase[purchase["Status"].ne("OK") | purchase["Warning"].fillna("").ne("")]
#                 if not problems.empty:
#                     sheets["Problems zakupu"] = problems
#                     self.log(f"У Реєстрі закупівлі знайдено проблемні рядки: {len(problems)}. Додано лист Problems zakupu.")
#                 self.log(f"Реєстр закупівлі готовий: {len(purchase)} рядків.")
#
#             if self.scan_pdf_var.get():
#                 self.log("Сканую PDF-реєстри у папці ‘Реєстри’…")
#                 pdf_index = scan_register_pdfs(paths.registers_dir)
#                 sheets["PDF rejestry"] = pdf_index
#                 self.log(f"Знайдено PDF-реєстрів з коректною назвою: {len(pdf_index)}.")
#
#             write_multi_sheet_excel_openpyxl(output, sheets)
#             self.log(f"ГОТОВО. Файл збережено: {output}")
#             messagebox.showinfo("Готово", f"Файл створено:\n{output}")
#         except Exception as exc:
#             self.log(f"ПОМИЛКА: {exc}")
#             messagebox.showerror("Помилка", str(exc))
#
#     def open_registers_dir(self) -> None:
#         try:
#             paths = ensure_project_structure(Path(self.project_dir.get()))
#             folder = str(paths.registers_dir)
#             if sys.platform == "darwin":
#                 subprocess.run(["open", folder], check=False)
#             elif sys.platform.startswith("win"):
#                 subprocess.run(["explorer", folder], check=False)
#             else:
#                 subprocess.run(["xdg-open", folder], check=False)
#         except Exception as exc:
#             self.log(f"Не вдалося відкрити папку: {exc}")
#
#
# def normalize_register_choice(value: Optional[str]) -> Optional[str]:
#     if value is None:
#         return None
#     value = value.strip().lower()
#     if value in ["1", "sales", "sprzedaz", "sprzedaż", "rejestr sprzedazy", "rejestr sprzedaży"]:
#         return "sales"
#     if value in ["2", "purchase", "zakup", "zakupy", "rejestr zakupu"]:
#         return "purchase"
#     raise ValueError(f"Unknown register type: {value}")
#
#
# def cli_main() -> None:
#     """Keeps old command-line mode available."""
#     parser = argparse.ArgumentParser(description="Generate Rejestr sprzedazy / zakupu from Iss-chem workbook, or open GUI.")
#     parser.add_argument("input", nargs="?", default=None, help="Optional path to source xlsx file. If omitted, uses <project>/accounting_table.xlsx")
#     parser.add_argument("--register", "-r", choices=["sales", "purchase", "1", "2", "both"], default=None)
#     parser.add_argument("--output", "-o", default=None, help="Output xlsx path")
#     parser.add_argument("--gui", action="store_true", help="Open GUI")
#     args = parser.parse_args()
#
#     if args.gui or (args.input is None and args.register is None and args.output is None):
#         root = Tk()
#         RejestrApp(root)
#         root.mainloop()
#         return
#
#     project_paths = ensure_project_structure(DEFAULT_PROJECT_DIR)
#     input_path = Path(args.input).expanduser().resolve() if args.input else project_paths.accounting_table
#     output_path = Path(args.output).expanduser().resolve() if args.output else project_paths.output_registers_workbook
#     register_type = args.register or "both"
#
#     sheets: dict[str, pd.DataFrame] = {}
#     if register_type in ["sales", "1", "both"]:
#         sales = generate_sales_register(input_path)
#         sheets["Rejestr sprzedazy"] = sales
#         sheets["Summary sprzedazy"] = build_sales_summary_df(sales)
#     if register_type in ["purchase", "2", "both"]:
#         purchase = generate_purchase_register(input_path)
#         sheets["Rejestr zakupu"] = purchase
#         sheets["Summary zakupu"] = build_purchase_summary_df(purchase)
#         problems = purchase[purchase["Status"].ne("OK") | purchase["Warning"].fillna("").ne("")]
#         if not problems.empty:
#             sheets["Problems zakupu"] = problems
#
#     write_multi_sheet_excel_openpyxl(output_path, sheets)
#     print(f"Created: {output_path}")
#
#
# if __name__ == "__main__":
#     cli_main()
