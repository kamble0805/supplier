"""
Extract all PO data (headers and line items) from all PDFs in ./pdfs and save to a single detailed sheet.
"""

import os
import re
import sys
import time
from pathlib import Path
import pandas as pd
import fitz  # PyMuPDF

BASE_DIR = Path(__file__).resolve().parent
PDF_FOLDER = BASE_DIR / "pdfs"
OUTPUT_EXCEL = BASE_DIR / "all_po_details.xlsx"

# ============================================================
# HELPERS
# ============================================================

def clean(value):
    if value is None:
        return ""
    value = str(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def first_match(patterns, text, flags=re.I | re.M):
    for pattern in patterns:
        m = re.search(pattern, text, flags)
        if m:
            return clean(m.group(1))
    return ""


def number(value):
    if value is None:
        return None
    value = str(value).replace(",", "").strip()
    m = re.search(r"-?\d+(?:\.\d+)?", value)
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


def get_pdf_text(pdf_path):
    """Extract text from all pages and reconstruct lines using coordinates."""
    doc = fitz.open(pdf_path)
    pages = []

    for page in doc:
        words = page.get_text("words")
        # Sort words by vertical coordinate y0
        words_sorted = sorted(words, key=lambda w: w[1])

        lines = []
        curr_line = []
        prev_y = None

        for w in words_sorted:
            y0 = w[1]
            if prev_y is None:
                curr_line.append(w)
            elif y0 - prev_y <= 2.0:
                curr_line.append(w)
            else:
                lines.append(curr_line)
                curr_line = [w]
            prev_y = y0
        if curr_line:
            lines.append(curr_line)

        page_text = []
        for line in lines:
            line_sorted = sorted(line, key=lambda w: w[0])
            page_text.append(" ".join(w[4] for w in line_sorted))

        pages.append("\n".join(page_text))

    doc.close()
    return "\n".join(pages)


# ============================================================
# PO HEADER EXTRACTION
# ============================================================

def clean_value(val):
    if not val:
        return ""
    val = re.split(r"\b(Total|CGST|SGST|IGST|PACKING|FREIGHT)\b", val, flags=re.I)[0]
    return clean(val)


def extract_header(text, pdf_name):
    data = {}

    m = re.search(r"^([A-Z0-9 &.,()'-]+)\s+([A-Z0-9]+/[A-Z0-9-]+/[A-Z0-9/]+)\s+(\d{2}-\d{2}-\d{4})$", text, re.M | re.I)
    if m:
        data["Supplier"] = clean_value(m.group(1))
        data["PO Number"] = clean_value(m.group(2))
        data["PO Date"] = clean_value(m.group(3))
    else:
        # Fallbacks
        data["Supplier"] = clean_value(first_match([r"PURCHASE ORDER NO\s*:\s*\n([^\n]+)"], text))
        data["PO Number"] = clean_value(first_match([r"\n([A-Z]{2,8}[A-Z0-9/.-]*\d{2,}[A-Z0-9/.-]*)\s+01-\d{2}-\d{4}"], text))
        data["PO Date"] = clean_value(first_match([r"(\d{2}-\d{2}-\d{4})"], text))

    # Requisition
    req_match = re.search(r"REQUISITION NO:\s*DATE\s*:\s*\n\s*([A-Z0-9 /-]+)\s+(\d{2}-\d{2}-\d{4})", text, re.I | re.M)
    if req_match:
        data["Requisition No"] = clean_value(req_match.group(1))
        data["Requisition Date"] = clean_value(req_match.group(2))
    else:
        data["Requisition No"] = clean_value(first_match([r"REQUISITION NO\s*:?\s*([^\n]+)"], text))
        data["Requisition Date"] = clean_value(first_match([r"REQUISITION NO\s*:?\s*[^\n]+\s+(\d{2}-\d{2}-\d{4})"], text))

    supplier_section = text
    if "SUPPLIER / MANUFACTURER" in text:
        supplier_section = text.split("SUPPLIER / MANUFACTURER", 1)[1]
    elif "PURCHASE ORDER" in text:
        supplier_section = text.split("PURCHASE ORDER", 1)[1]

    data["Supplier GST"] = clean_value(first_match([r"GST NO[ \t]*(?::[ \t]*)?([0-9A-Z]{15})"], supplier_section))
    data["Supplier PAN"] = clean_value(first_match([r"PAN NO[ \t]*(?::[ \t]*)?([A-Z0-9]{10})"], supplier_section))
    data["Supplier Email"] = clean_value(first_match([r"E-MAIL[ \t]*(?::[ \t]*)?([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})"], supplier_section))
    data["Supplier Contact"] = clean_value(first_match([r"CONTACT NO[ \t]*(?::[ \t]*)?([a-zA-Z0-9].*)"], supplier_section))
    data["Payment Terms"] = clean_value(first_match([r"PAYMENT TERMS[ \t]*(?::[ \t]*)?([a-zA-Z0-9].*)"], supplier_section))
    data["Delivery Schedule"] = clean_value(first_match([r"DELIVERY SCHEDULE[ \t]*(?::[ \t]*)?([a-zA-Z0-9].*)"], supplier_section))
    data["Delivery Mode"] = clean_value(first_match([r"DELIVERY MODE[ \t]*(?::[ \t]*)?([a-zA-Z0-9].*)"], supplier_section))

    # Address Lines
    address_lines = []
    lines = [clean(l) for l in text.splitlines() if clean(l)]
    supplier_found = False
    for line in lines:
        if not supplier_found:
            if data.get("Supplier") and data["Supplier"].lower() in line.lower():
                supplier_found = True
            continue
        if data.get("PO Number") and data["PO Number"].lower() in line.lower():
            continue
        if data.get("PO Date") and data["PO Date"].lower() in line.lower():
            continue
        stop_keywords = ["REQUISITION", "STATE", "PAN NO", "GST NO", "E-MAIL", "CONTACT NO", "SR NO", "DESCRIPTION", "YOUR REF NO"]
        if any(kw in line.upper() for kw in stop_keywords):
            for kw in stop_keywords:
                if kw in line.upper():
                    idx = line.upper().index(kw)
                    part = clean_value(line[:idx])
                    if part:
                        address_lines.append(part)
                    break
            break
        address_lines.append(clean_value(line))
        if len(address_lines) >= 3:
            break

    data["Add_Line1"] = address_lines[0] if len(address_lines) > 0 else ""
    data["Add_Line2"] = address_lines[1] if len(address_lines) > 1 else ""
    data["Add_Line3"] = address_lines[2] if len(address_lines) > 2 else ""

    full_address = " ".join(address_lines)
    pin_match = re.search(r"\b(\d{6})\b", full_address)
    data["PinCode"] = pin_match.group(1) if pin_match else ""

    city_match = re.search(r"\b([A-Za-z\s]+?)(?:\s*-\s*|\s+)\d{6}", full_address)
    data["City"] = clean(city_match.group(1)) if city_match else ""

    data["StateName"] = clean_value(first_match([r"STATE[ \t]*(?::[ \t]*)?([a-zA-Z0-9\s].*)"], supplier_section))
    if not data["StateName"] and "MAHARASHTRA" in full_address.upper():
        data["StateName"] = "MAHARASHTRA"

    data["CountryCode"] = "IN" if "INDIA" in full_address.upper() or "MAHARASHTRA" in full_address.upper() or data["StateName"] else ""
    data["File Name"] = pdf_name

    return data


# ============================================================
# LINE ITEM EXTRACTION
# ============================================================

def extract_items(text):
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    items = []

    row_re = re.compile(
        r"^(?P<sr>\d+)\s+"
        r"(?P<description>.+?)\s+"
        r"(?P<qty>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<unit>[A-Za-z]+)\s+"
        r"(?P<rate>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<rate_unit>[A-Za-z /]+)\s+"
        r"(?P<value>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<discount_amt>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<discount_pct>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<after_discount>[\d,]+(?:\.\d+)?)$",
        re.I
    )

    row_re_alt = re.compile(
        r"^(?P<sr>\d+)\s+"
        r"(?P<description>.+?)\s+"
        r"(?P<qty>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<unit>[A-Za-z]+)\s+"
        r"(?P<rate>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<value>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<discount_amt>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<discount_pct>[\d,]+(?:\.\d+)?)\s+"
        r"(?P<after_discount>[\d,]+(?:\.\d+)?)$",
        re.I
    )

    current_item = None
    stop_pattern = re.compile(r"^(Total|PACKING|FREIGHT|CGST|SGST|IGST|DELIVERY|INSPECTION|PAYMENT|SPECIAL|NOTES|For,)", re.I)

    for line in lines:
        m = row_re.match(line) or row_re_alt.match(line)
        if m:
            if current_item:
                items.append(current_item)
            current_item = m.groupdict()
            current_item["description"] = clean(current_item["description"])
        elif current_item:
            if stop_pattern.match(line):
                items.append(current_item)
                current_item = None
            else:
                current_item["description"] = clean(current_item["description"] + " " + line)

    if current_item:
        items.append(current_item)

    return items


def extract_totals(text):
    return {
        "CGST": first_match([r"CGST\s+.*?([\d,]+(?:\.\d+)?)$"], text),
        "SGST": first_match([r"SGST\s+.*?([\d,]+(?:\.\d+)?)$"], text),
        "IGST": first_match([r"IGST\s+.*?([\d,]+(?:\.\d+)?)$"], text),
        "Grand Total": first_match([r"Total\s*:\s*([\d,]+(?:\.\d+)?)"], text)
    }


# ============================================================
# MAIN EXECUTION
# ============================================================

def main():
    print("=" * 60)
    print("EXTRACT ALL PO DETAILS (HEADERS + LINE ITEMS)")
    print("=" * 60)

    if not PDF_FOLDER.exists():
        raise FileNotFoundError(f"PDF folder not found: {PDF_FOLDER}")

    pdf_files = sorted(PDF_FOLDER.glob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in: {PDF_FOLDER}")

    print(f"Found {len(pdf_files)} PDF files to process.")

    output_rows = []

    for idx, pdf_path in enumerate(pdf_files, start=1):
        print(f"[{idx}/{len(pdf_files)}] Processing: {pdf_path.name}")
        try:
            text = get_pdf_text(pdf_path)
            # Strip terms and conditions boilerplate
            text = re.split(r"TERMS\s+(?:AND|&)\s+CONDITIONS", text, flags=re.I)[0]
            header_data = extract_header(text, pdf_path.name)
            totals = extract_totals(text)
            items = extract_items(text)

            if not items:
                # Still add one row with header info if no items were extracted
                items = [{}]

            for item in items:
                row_data = {
                    "Supplier Name": header_data.get("Supplier", ""),
                    "PO Number": header_data.get("PO Number", ""),
                    "PO Date": header_data.get("PO Date", ""),
                    "Add_Line1": header_data.get("Add_Line1", ""),
                    "Add_Line2": header_data.get("Add_Line2", ""),
                    "Add_Line3": header_data.get("Add_Line3", ""),
                    "City": header_data.get("City", ""),
                    "StateName": header_data.get("StateName", ""),
                    "PinCode": header_data.get("PinCode", ""),
                    "CountryCode": header_data.get("CountryCode", ""),
                    "Supplier GST": header_data.get("Supplier GST", ""),
                    "Supplier PAN": header_data.get("Supplier PAN", ""),
                    "Supplier Email": header_data.get("Supplier Email", ""),
                    "Supplier Contact": header_data.get("Supplier Contact", ""),
                    "Payment Terms": header_data.get("Payment Terms", ""),
                    "Delivery Schedule": header_data.get("Delivery Schedule", ""),
                    "Delivery Mode": header_data.get("Delivery Mode", ""),
                    "Requisition No": header_data.get("Requisition No", ""),
                    "Requisition Date": header_data.get("Requisition Date", ""),
                    "Line Sr No": item.get("sr", ""),
                    "Item Description": item.get("description", ""),
                    "Qty": number(item.get("qty")),
                    "Unit": item.get("unit", ""),
                    "Rate": number(item.get("rate")),
                    "Rate Unit": item.get("rate_unit", ""),
                    "Value": number(item.get("value")),
                    "Discount Amt": number(item.get("discount_amt")),
                    "Discount Pct": number(item.get("discount_pct")),
                    "After Discount": number(item.get("after_discount")),
                    "CGST": number(totals.get("CGST")),
                    "SGST": number(totals.get("SGST")),
                    "IGST": number(totals.get("IGST")),
                    "Grand Total": number(totals.get("Grand Total")),
                    "File Name": pdf_path.name
                }
                output_rows.append(row_data)

        except Exception as e:
            print(f"Error processing {pdf_path.name}: {e}")

    df = pd.DataFrame(output_rows)

    # Save logic with PermissionError checking
    target_excel = OUTPUT_EXCEL
    for i in range(100):
        try:
            if os.path.exists(target_excel):
                with open(target_excel, 'a'):
                    pass
            break
        except PermissionError:
            target_excel = OUTPUT_EXCEL.parent / f"{OUTPUT_EXCEL.stem}_{i+1}{OUTPUT_EXCEL.suffix}"

    if target_excel != OUTPUT_EXCEL:
        print(f"\nWARNING: '{OUTPUT_EXCEL.name}' is open or locked in Excel.")
        print(f"Saving to alternative file: '{target_excel.name}'")

    df.to_excel(target_excel, index=False)
    print(f"\nCompleted! Saved all extracted PO details to: {target_excel}")
    print(f"Total Rows Extracted: {len(df)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
