from __future__ import annotations

import re
from pathlib import Path

import fitz
import pandas as pd
import pytesseract
from PIL import Image

from core.numbers import float_to_pl_text, pl_number_to_float


def extract_amounts(line: str) -> list[str]:
    return re.findall(r"-?\d[\d\s\u00a0]*[,.]\d{2}", str(line))


def normalize_date(value: str) -> str:
    match = re.search(r"(\d{2})[./](\d{2})[./](\d{4})", str(value))
    if not match:
        return ""
    day, month, year = match.groups()
    return f"{day}.{month}.{year}"


def get_first_date(value: str) -> str:
    matches = re.findall(r"\d{2}[./]\d{2}[./]\d{4}", str(value))
    return normalize_date(matches[0]) if matches else ""


def get_last_date(value: str) -> str:
    matches = re.findall(r"\d{2}[./]\d{2}[./]\d{4}", str(value))
    return normalize_date(matches[-1]) if matches else ""


def looks_like_summary_line(line: str) -> bool:
    lower = str(line).lower()
    return any(word in lower for word in ["razem", "suma", "podsumowanie", "łącznie", "lacznie", "total"])


def detect_register_type(text: str) -> str:
    text_upper = str(text).upper()
    if re.search(r"\bZAKUP\b", text_upper):
        return "ZAKUP"
    if re.search(r"\bSPRZEDA[ŻZ]\b", text_upper):
        return "SPRZEDAŻ"
    return ""


def extract_accounting_id(text: str) -> str:
    match = re.search(r"\b\d+/\d+/(SPRZEDA[ŻZ]|ZAKUP)\b", str(text), flags=re.IGNORECASE)
    return match.group(0) if match else ""


def extract_register_name(text: str) -> str:
    match = re.search(r"\b(SPRZEDA[ŻZ]|ZAKUP)\b", str(text), flags=re.IGNORECASE)
    if not match:
        return ""
    value = match.group(0).upper()
    return "SPRZEDAŻ" if value in {"SPRZEDAZ", "SPRZEDAŻ"} else value


def extract_rate(text: str) -> str:
    match = re.search(r"\b(NP|\d{1,2}%)\b", str(text), flags=re.IGNORECASE)
    return match.group(0).upper() if match else ""


def split_contractor(value: str) -> tuple[str, str]:
    if not value:
        return "", ""

    value = str(value).strip()
    for source, target in {
        "Połtawa": "Poltawa",
        "Poftawa": "Poltawa",
        "Poitawa": "Poltawa",
        "Pottawa": "Poltawa",
    }.items():
        value = value.replace(source, target)

    company_endings = [
        "LLC",
        "LTD",
        "Ltd",
        "B.V.",
        "BV",
        "Sp. z o.o.",
        "SP. Z O.O.",
        "S.A.",
        "SA",
        "ODPOWIEDZIALNOSCIA",
        "ODPOWIEDZIALNOŚCIĄ",
    ]
    for ending in company_endings:
        match = re.search(rf"({re.escape(ending)})[\.,]\s*(.+)$", value)
        if match:
            return value[: match.end(1)].strip(), match.group(2).strip(" .,")

    if "," in value:
        name, address = value.split(",", 1)
        return name.strip(), address.strip(" .,")

    city_match = re.search(r"\.\s*(Poltawa.*)$", value, flags=re.IGNORECASE)
    if city_match:
        return value[: city_match.start()].strip(), city_match.group(1).strip(" .,")

    return value, ""


def clean_lp(value: str) -> str:
    replacements = {"l": "1", "I": "1", "|": "1", "i": "1", "O": "0", "o": "0"}
    cleaned = "".join(replacements.get(char, char) for char in str(value).strip())
    digits = re.findall(r"\d+", cleaned)
    return digits[0] if digits else ""


def pdf_to_images(pdf_path: Path, dpi: int = 300) -> list[Image.Image]:
    images = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            if pix.alpha:
                image = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples).convert("RGB")
            else:
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(image)
    return images


def ocr_image(image: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(image, lang="eng+pol")
    except pytesseract.TesseractError:
        return pytesseract.image_to_string(image, lang="eng")


def find_pdf_summary(all_ocr_lines: list[dict]) -> tuple[dict | None, pd.DataFrame]:
    candidates = []
    for item in all_ocr_lines:
        amounts = extract_amounts(item["Line"])
        if len(amounts) >= 3 and looks_like_summary_line(item["Line"]):
            candidates.append(
                {
                    "Page": item["Page"],
                    "Raw line": item["Line"],
                    "Netto": float_to_pl_text(pl_number_to_float(amounts[-3])),
                    "VAT": float_to_pl_text(pl_number_to_float(amounts[-2])),
                    "Brutto": float_to_pl_text(pl_number_to_float(amounts[-1])),
                    "Netto_num": pl_number_to_float(amounts[-3]),
                    "VAT_num": pl_number_to_float(amounts[-2]),
                    "Brutto_num": pl_number_to_float(amounts[-1]),
                }
            )
    if not candidates:
        return None, pd.DataFrame()
    return candidates[-1], pd.DataFrame(candidates)


def parse_record_from_two_lines(line1: str, line2: str, page: int) -> dict:
    amounts = extract_amounts(line1)
    netto_raw, vat_raw, brutto_raw = (amounts[-3], amounts[-2], amounts[-1]) if len(amounts) >= 3 else ("", "", "")
    netto_num = pl_number_to_float(netto_raw)
    vat_num = pl_number_to_float(vat_raw)
    brutto_num = pl_number_to_float(brutto_raw)

    clean_line1 = line1
    for amount in amounts[-3:]:
        clean_line1 = clean_line1.replace(amount, "", 1)

    data1 = get_first_date(clean_line1)
    data2 = get_last_date(line2)
    typ_rejestru = detect_register_type(clean_line1)

    lp_match = re.match(r"\s*([0-9lLiI|oO]+)\b", clean_line1)
    lp = clean_lp(lp_match.group(1)) if lp_match else ""

    nr_w_rej_match = re.match(r"\s*(\d+)\b", line2)
    nr_w_rej = nr_w_rej_match.group(1) if nr_w_rej_match else ""

    nr_dokumentu = ""
    if nr_w_rej:
        tmp = re.sub(r"^\s*" + re.escape(nr_w_rej) + r"\b", "", line2).strip()
        nr_dokumentu = tmp.split(data2, 1)[0].strip() if data2 and data2 in tmp else tmp.strip()

    nip_pesel = ""
    if data1 and data1 in clean_line1:
        after_date = clean_line1.split(data1, 1)[1].strip()
        nip_match = re.match(r"([A-Z]{2}[A-Z0-9]*|\d{6,})\b", after_date)
        nip_pesel = nip_match.group(1) if nip_match else ""

    stawka = extract_rate(clean_line1)
    contractor_text = clean_line1
    if nip_pesel and nip_pesel in contractor_text:
        contractor_text = contractor_text.split(nip_pesel, 1)[1].strip()
    elif data1 and data1 in contractor_text:
        contractor_text = contractor_text.split(data1, 1)[1].strip()
    if stawka and stawka in contractor_text:
        contractor_text = contractor_text.split(stawka, 1)[0].strip()

    contractor_name, contractor_address = split_contractor(contractor_text.strip(" -|"))
    return {
        "Page": page,
        "Typ rejestru": typ_rejestru,
        "Lp": lp,
        "Nr w rej.": nr_w_rej,
        "Id księgowy": extract_accounting_id(clean_line1),
        "Nr dokumentu": nr_dokumentu,
        "Rejestr": extract_register_name(clean_line1),
        "Korekta": re.search(r"\bKOREKTA\b", clean_line1, flags=re.IGNORECASE).group(0)
        if re.search(r"\bKOREKTA\b", clean_line1, flags=re.IGNORECASE)
        else "",
        "Data 1": data1,
        "Data 2": data2,
        "NIP/PESEL": nip_pesel,
        "Kontrahent - nazwa": contractor_name,
        "Kontrahent - adres": contractor_address,
        "Stawka": stawka,
        "Netto": float_to_pl_text(netto_num),
        "VAT": float_to_pl_text(vat_num),
        "Brutto": float_to_pl_text(brutto_num),
        "Netto_num": netto_num,
        "VAT_num": vat_num,
        "Brutto_num": brutto_num,
        "Raw Netto OCR": netto_raw,
        "Raw VAT OCR": vat_raw,
        "Raw Brutto OCR": brutto_raw,
        "Raw line 1": line1,
        "Raw line 2": line2,
        "Review_note": "",
    }


def parse_record_from_one_line(line: str, page: int) -> dict:
    parsed = parse_record_from_two_lines(line, "", page)
    parsed["Nr w rej."] = ""
    parsed["Nr dokumentu"] = ""
    parsed["Data 2"] = ""
    parsed["Raw line 2"] = ""
    parsed["Review_note"] = "Запис розпізнано як один рядок — перевір Nr w rej., Nr dokumentu, Data 2."
    return parsed


def build_rows_from_ocr_lines(all_ocr_lines: list[dict]) -> list[dict]:
    rows = []
    lines_by_page: dict[int, list[str]] = {}
    for item in all_ocr_lines:
        lines_by_page.setdefault(item["Page"], []).append(item["Line"])

    for page, lines in lines_by_page.items():
        i = 0
        while i < len(lines):
            line1 = lines[i]
            if looks_like_summary_line(line1):
                i += 1
                continue

            has_lp = bool(re.match(r"\s*[0-9lLiI|oO]+\b", line1))
            has_id = bool(extract_accounting_id(line1))
            has_date = bool(re.search(r"\d{2}[./]\d{2}[./]\d{4}", line1))
            has_amounts = len(extract_amounts(line1)) >= 3

            if has_lp and has_id and has_date and has_amounts:
                line2 = ""
                if i + 1 < len(lines):
                    possible_line2 = lines[i + 1]
                    possible_has_nr = bool(re.match(r"\s*\d+\b", possible_line2))
                    possible_has_date = bool(re.search(r"\d{2}[./]\d{2}[./]\d{4}", possible_line2))
                    if possible_has_nr or possible_has_date:
                        line2 = possible_line2
                        i += 1
                rows.append(parse_record_from_two_lines(line1, line2, page) if line2 else parse_record_from_one_line(line1, page))
            i += 1

    return rows


def add_semantic_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    if "Typ rejestru" not in df.columns:
        return df
    df["Data wpł."] = pd.NaT
    df["Data wyst. zakupu"] = pd.NaT
    df["Data wyst. sprzedaży"] = pd.NaT
    df["Data sprz."] = pd.NaT
    for idx, row in df.iterrows():
        if row.get("Typ rejestru") == "ZAKUP":
            df.at[idx, "Data wpł."] = row.get("Data 1")
            df.at[idx, "Data wyst. zakupu"] = row.get("Data 2")
        elif row.get("Typ rejestru") == "SPRZEDAŻ":
            df.at[idx, "Data wyst. sprzedaży"] = row.get("Data 1")
            df.at[idx, "Data sprz."] = row.get("Data 2")
    return df


def add_review_notes(df: pd.DataFrame) -> pd.DataFrame:
    for idx, row in df.iterrows():
        notes = []
        existing_note = row.get("Review_note", "")
        if isinstance(existing_note, str) and existing_note:
            notes.append(existing_note)
        for field in [
            "Typ rejestru",
            "Lp",
            "Nr w rej.",
            "Id księgowy",
            "Nr dokumentu",
            "Data 1",
            "Data 2",
            "Kontrahent - nazwa",
            "Netto_num",
            "VAT_num",
            "Brutto_num",
        ]:
            value = row.get(field, "")
            if pd.isna(value) or value == "":
                notes.append(f"Перевірити {field}")

        vat = row.get("VAT_num")
        netto = row.get("Netto_num")
        brutto = row.get("Brutto_num")
        if not pd.isna(vat) and not pd.isna(netto) and not pd.isna(brutto):
            if abs(vat) <= 0.01 and abs(netto - brutto) > 0.01:
                notes.append("Netto не дорівнює Brutto при VAT = 0")
            elif abs(vat) > 0.01 and abs((netto + vat) - brutto) > 0.02:
                notes.append("Netto + VAT не дорівнює Brutto")
        df.at[idx, "Review_note"] = "; ".join(dict.fromkeys(notes))
    return df


def fix_amount_mismatches(df: pd.DataFrame) -> pd.DataFrame:
    for idx, row in df.iterrows():
        netto = row.get("Netto_num")
        vat = row.get("VAT_num")
        brutto = row.get("Brutto_num")
        if pd.isna(netto) or pd.isna(vat) or pd.isna(brutto):
            continue
        if abs(vat) <= 0.01 and abs(netto - brutto) > 0.01:
            old_netto = netto
            df.at[idx, "Netto_num"] = brutto
            df.at[idx, "Netto"] = float_to_pl_text(brutto)
            note = row.get("Review_note", "")
            correction_note = (
                f"Netto виправлено автоматично: було {float_to_pl_text(old_netto)}, "
                f"стало {float_to_pl_text(brutto)} на основі Brutto, бо VAT = 0"
            )
            df.at[idx, "Review_note"] = note + "; " + correction_note if note else correction_note
    return df


def normalize_amount_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col_text, col_num in [("Netto", "Netto_num"), ("VAT", "VAT_num"), ("Brutto", "Brutto_num")]:
        if col_text in df.columns and col_num in df.columns:
            df[col_text] = df[col_num].apply(float_to_pl_text)
    return df


def normalize_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["Data 1", "Data 2"]:
        if col in df.columns:
            df[col] = df[col].apply(normalize_date)
            df[col] = pd.to_datetime(df[col], format="%d.%m.%Y", errors="coerce")
    return df


def extract_pdf_register(pdf_path: Path, log_func=None) -> dict[str, pd.DataFrame]:
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF не знайдено: {pdf_path}")

    all_ocr_lines = []
    pages = pdf_to_images(pdf_path, dpi=300)
    for page_num, image in enumerate(pages, start=1):
        if log_func:
            log_func(f"OCR PDF, сторінка {page_num}/{len(pages)}")
        text = ocr_image(image)
        for line in text.splitlines():
            line = line.strip()
            if line:
                all_ocr_lines.append({"Page": page_num, "Line": line})

    rows_df = pd.DataFrame(build_rows_from_ocr_lines(all_ocr_lines))
    if not rows_df.empty:
        rows_df = add_review_notes(rows_df)
        rows_df = fix_amount_mismatches(rows_df)
        rows_df = normalize_amount_text_columns(rows_df)
        rows_df = normalize_date_columns(rows_df)
        rows_df = add_semantic_date_columns(rows_df)

    pdf_summary, summary_candidates_df = find_pdf_summary(all_ocr_lines)
    check_df = pd.DataFrame(
        {
            "Metric": ["Netto total", "VAT total", "Brutto total"],
            "Calculated from rows": [
                rows_df["Netto_num"].sum() if "Netto_num" in rows_df else None,
                rows_df["VAT_num"].sum() if "VAT_num" in rows_df else None,
                rows_df["Brutto_num"].sum() if "Brutto_num" in rows_df else None,
            ],
            "PDF summary found": [
                pdf_summary["Netto_num"] if pdf_summary else None,
                pdf_summary["VAT_num"] if pdf_summary else None,
                pdf_summary["Brutto_num"] if pdf_summary else None,
            ],
            "Summary source": [
                f"Знайдено автоматично, сторінка {pdf_summary['Page']}" if pdf_summary else "Не знайдено автоматично"
            ]
            * 3,
        }
    )
    check_df["Difference"] = check_df["Calculated from rows"] - check_df["PDF summary found"]

    all_lines_df = pd.DataFrame(all_ocr_lines)
    return {
        "OCR_rows": rows_df,
        "All_OCR_lines": all_lines_df,
        "Check": check_df,
        "Summary_candidates": summary_candidates_df,
    }
