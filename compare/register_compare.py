from pathlib import Path
from datetime import date
import re
import pandas as pd

from core.numbers import clean_doc_number
from core.paths import ProjectPaths, parse_register_pdf_filename
from core.excel_format import write_multi_sheet_excel_openpyxl
from pdf_tools.pdf_register import extract_pdf_register
from registers.our_registers import generate_sales_register, generate_purchase_register

def _cmp_doc(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).upper().strip()
    text = re.sub(r"\s+", "", text)
    text = text.replace("SPRZEDAŻ", "SPRZEDAZ")
    return text

def _date_in_period(series: pd.Series, start_date: date, end_date: date) -> pd.Series:
    d = pd.to_datetime(series, errors="coerce").dt.date
    return d.ge(start_date) & d.le(end_date)

def build_our_register_for_pdf_period(accounting_table: Path, register_kind: str, start_date: date, end_date: date) -> pd.DataFrame:
    if register_kind == "purchase":
        df = generate_purchase_register(accounting_table)
        df = df[_date_in_period(df["Date of issue"], start_date, end_date)].copy()
        df["Compare document"] = df["Invoice number"].apply(_cmp_doc)
        df["Compare date"] = pd.to_datetime(df["Date of issue"], errors="coerce").dt.date
        df["Compare contractor"] = df["Supplier / Company"]
    else:
        df = generate_sales_register(accounting_table)
        df = df[_date_in_period(df["Date of issue"], start_date, end_date)].copy()
        df["Compare document"] = df["Nr dokumentu"].apply(_cmp_doc)
        df["Compare date"] = pd.to_datetime(df["Date of issue"], errors="coerce").dt.date
        df["Compare contractor"] = df["Customer"]
    return df

def prepare_pdf_compare_df(pdf_rows: pd.DataFrame, register_kind: str, start_date: date, end_date: date) -> pd.DataFrame:
    df = pdf_rows.copy()
    expected_type = "ZAKUP" if register_kind == "purchase" else "SPRZEDAŻ"
    if "Typ rejestru" in df.columns:
        typed = df[df["Typ rejestru"].eq(expected_type)].copy()
        if not typed.empty:
            df = typed
    date_col = "Data wyst. zakupu" if register_kind == "purchase" else "Data wyst. sprzedaży"
    if date_col not in df.columns or df[date_col].isna().all():
        date_col = "Data 2" if register_kind == "purchase" else "Data 1"
    df = df[_date_in_period(df[date_col], start_date, end_date)].copy()
    df["Compare document"] = df["Nr dokumentu"].apply(_cmp_doc)
    df["Compare date"] = pd.to_datetime(df[date_col], errors="coerce").dt.date
    df["Compare contractor"] = df.get("Kontrahent - nazwa", "")
    return df

def compare_registers(our_df: pd.DataFrame, pdf_df: pd.DataFrame, register_kind: str) -> dict[str, pd.DataFrame]:
    our_doc_col = "Invoice number" if register_kind == "purchase" else "Nr dokumentu"
    our_base_cols = ["Compare document", "Compare date", "Compare contractor", our_doc_col, "Netto PLN", "VAT PLN", "Brutto PLN"]
    pdf_base_cols = ["Compare document", "Compare date", "Compare contractor", "Nr dokumentu", "Netto_num", "VAT_num", "Brutto_num", "Review_note"]
    our = our_df[[c for c in our_base_cols if c in our_df.columns]].copy()
    pdf = pdf_df[[c for c in pdf_base_cols if c in pdf_df.columns]].copy()
    # Aggregate by document number because OCR or source data may contain repeated rows/rates.
    our_g = our.groupby("Compare document", dropna=False).agg({
        "Compare date": "first", "Compare contractor": "first", our_doc_col: "first",
        "Netto PLN": "sum", "VAT PLN": "sum", "Brutto PLN": "sum"
    }).reset_index()
    pdf_g = pdf.groupby("Compare document", dropna=False).agg({
        "Compare date": "first", "Compare contractor": "first", "Nr dokumentu": "first",
        "Netto_num": "sum", "VAT_num": "sum", "Brutto_num": "sum", "Review_note": lambda x: "; ".join([str(v) for v in x if str(v)])
    }).reset_index()
    merged = our_g.merge(pdf_g, on="Compare document", how="outer", suffixes=(" ours", " pdf"), indicator=True)
    statuses = []
    notes = []
    for _, row in merged.iterrows():
        if row["_merge"] == "left_only":
            statuses.append("Тільки у нас")
            notes.append("Документ є у нашому accounting_table, але не знайдений у PDF")
            continue
        if row["_merge"] == "right_only":
            statuses.append("Тільки в PDF аутсорсингу")
            notes.append("Документ є у PDF, але не знайдений у нашому accounting_table")
            continue
        diffs = []
        for label, ours_col, pdf_col in [("Netto", "Netto PLN", "Netto_num"), ("VAT", "VAT PLN", "VAT_num"), ("Brutto", "Brutto PLN", "Brutto_num")]:
            ours_val = row.get(ours_col)
            pdf_val = row.get(pdf_col)
            if pd.notna(ours_val) and pd.notna(pdf_val) and abs(float(ours_val) - float(pdf_val)) > 0.02:
                diffs.append(f"{label}: наші {float_to_pl_text(ours_val)} vs PDF {float_to_pl_text(pdf_val)}")
        od, pdte = row.get("Compare date ours"), row.get("Compare date pdf")
        if pd.notna(od) and pd.notna(pdte) and od != pdte:
            diffs.append(f"Дата: наші {od} vs PDF {pdte}")
        if diffs:
            statuses.append("Різниця")
            notes.append("; ".join(diffs))
        else:
            statuses.append("OK")
            notes.append("")
    merged.insert(0, "Status", statuses)
    merged.insert(1, "Difference note", notes)
    only_ours = merged[merged["Status"].eq("Тільки у нас")].copy()
    only_pdf = merged[merged["Status"].eq("Тільки в PDF аутсорсингу")].copy()
    differences = merged[merged["Status"].eq("Різниця")].copy()
    ok = merged[merged["Status"].eq("OK")].copy()
    return {"Порівняння": merged, "OK": ok, "Тільки у нас": only_ours, "Тільки в PDF": only_pdf, "Різниці": differences}

def comparison_output_path(paths: ProjectPaths, pdf_path: Path, parsed: dict) -> Path:
    kind = "Закупівлі" if parsed["Register type"] == "Закупівлі" else "Продажу"
    start = parsed["Period start"].strftime("%Y-%m-%d")
    end = parsed["Period end"].strftime("%Y-%m-%d")
    stem = re.sub(r"[^A-Za-zА-Яа-яІіЇїЄєҐґ0-9_. -]+", "_", pdf_path.stem)
    return paths.accounting_report_dir / f"Порівняння {kind} {start} - {end} ({stem}).xlsx"

def create_pdf_comparison_workbook(paths: ProjectPaths, pdf_path: Path, log_func=None) -> Path:
    parsed = parse_register_pdf_filename(pdf_path)
    if not parsed:
        raise ValueError(f"Назва PDF не відповідає шаблону Реєстр Закупівлі/Продажу MMYYYY-MMYYYY або DDMMYYYY-DDMMYYYY: {pdf_path.name}")
    register_kind = "purchase" if parsed["Register type"] == "Закупівлі" else "sales"
    start_date = parsed["Period start"]
    end_date = parsed["Period end"]
    if log_func:
        log_func(f"Період PDF: {start_date} — {end_date}; тип: {parsed['Register type']}")
    pdf_sheets = extract_pdf_register(pdf_path, log_func=log_func)
    pdf_rows = prepare_pdf_compare_df(pdf_sheets["OCR_rows"], register_kind, start_date, end_date)
    our_rows = build_our_register_for_pdf_period(paths.accounting_table, register_kind, start_date, end_date)
    cmp_sheets = compare_registers(our_rows, pdf_rows, register_kind)
    meta = pd.DataFrame([{
        "PDF file": str(pdf_path), "Register type": parsed["Register type"],
        "Period start": start_date, "Period end": end_date,
        "Our rows in period": len(our_rows), "PDF rows in period": len(pdf_rows),
    }])
    sheets = {
        "Info": meta,
        "Наш реєстр": our_rows,
        "Реєстр PDF": pdf_rows,
        **cmp_sheets,
        "PDF Check": pdf_sheets["Check"],
        "PDF OCR all lines": pdf_sheets["All_OCR_lines"],
    }
    output = comparison_output_path(paths, pdf_path, parsed)
    write_multi_sheet_excel_openpyxl(output, sheets)
    return output
