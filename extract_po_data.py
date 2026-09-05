"""
PDF Purchase Order -> Excel extraction utility

Folder structure:
project/
  extract_po_data.py
  template.xlsx
  pdfs/
      01-08-2026.pdf
      01-08-2026 2.pdf
      ...

Run:
    pip install pandas openpyxl pymupdf
    python extract_po_data.py

The script:
1. Opens template.xlsx.
2. Reads Sheet1 headers.
3. Creates/replaces a new sheet named "PDF Extracted Data".
4. Extracts PO header + line-item information from every PDF in ./pdfs.
5. Uses the EXACT same columns as Sheet1.
6. Saves the result as template_PDF_Extracted.xlsx.

Important:
- If a column cannot be identified in a PDF, it is left blank.
- The original Sheet1 data and formatting are retained.
- Multiple line items in one PO become multiple Excel rows.
"""

import re
from pathlib import Path
from copy import copy
import pandas as pd
import fitz  # PyMuPDF
from openpyxl import load_workbook

# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
INPUT_EXCEL = BASE_DIR / "template.xlsx"
PDF_FOLDER = BASE_DIR / "pdfs"
OUTPUT_EXCEL = BASE_DIR / "template_PDF_Extracted.xlsx"

SOURCE_SHEET = "Sheet1"
OUTPUT_SHEET = "PDF Extracted Data"

# ============================================================
# HELPERS
# ============================================================

def clean(value):
    if value is None:
        return ""
    value = str(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", clean(value).lower())


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


def first_match(patterns, text, flags=re.I | re.M):
    for pattern in patterns:
        m = re.search(pattern, text, flags)
        if m:
            return clean(m.group(1))
    return ""


def number(value):
    """Convert Indian-style/simple numeric text to float where possible."""
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


# ============================================================
# PO HEADER EXTRACTION
# ============================================================

def clean_value(val):
    if not val:
        return ""
    # Strip any trailing taxes/totals like 'Total : ...', 'FREIGHT ...', etc.
    val = re.split(r"\b(?:Total|CGST|SGST|IGST|PACKING|FREIGHT)\b", val, flags=re.I)[0]
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

    # Search for supplier metadata only in the supplier section to avoid buyer info at top
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

    # Extract Address Lines
    address_lines = []
    lines = [clean(l) for l in text.splitlines() if clean(l)]
    supplier_found = False
    for line in lines:
        if not supplier_found:
            if data.get("Supplier") and data["Supplier"].lower() in line.lower():
                supplier_found = True
            continue
        # Skip PO Number and Date line if it appears before address
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

    # State Name
    data["StateName"] = clean_value(first_match([r"STATE[ \t]*(?::[ \t]*)?([a-zA-Z0-9\s].*)"], supplier_section))
    if not data["StateName"] and "MAHARASHTRA" in full_address.upper():
        data["StateName"] = "MAHARASHTRA"

    # Country Code
    data["CountryCode"] = "IN" if "INDIA" in full_address.upper() or "MAHARASHTRA" in full_address.upper() or data["StateName"] else ""
    data["File Name"] = pdf_name

    return data


# ============================================================
# LINE ITEM EXTRACTION
# ============================================================

def extract_items(text):
    """
    Extract standard PO rows from coordinate-sorted text.
    """
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
            # Check if this is a stop line
            if stop_pattern.match(line):
                items.append(current_item)
                current_item = None
            else:
                # Append to active item's description
                current_item["description"] = clean(current_item["description"] + " " + line)

    if current_item:
        items.append(current_item)

    return items


# ============================================================
# TAX / TOTAL EXTRACTION
# ============================================================

def extract_totals(text):
    return {
        "CGST": first_match([r"CGST\s+.*?([\d,]+(?:\.\d+)?)$"], text),
        "SGST": first_match([r"SGST\s+.*?([\d,]+(?:\.\d+)?)$"], text),
        "IGST": first_match([r"IGST\s+.*?([\d,]+(?:\.\d+)?)$"], text),
        "Grand Total": first_match([r"Total\s*:\s*([\d,]+(?:\.\d+)?)"], text)
    }


# ============================================================
# MAP EXTRACTED DATA TO SHEET1 COLUMNS
# ============================================================

COLUMN_ALIASES = {
    "srno": "sr",
    "serialno": "sr",
    "itemno": "item_no",
    "itemcode": "item_no",
    "partcode": "item_no",
    "partno": "item_no",
    "description": "description",
    "itemdescription": "description",
    "itemdesc": "description",
    "qty": "qty",
    "quantity": "qty",
    "qtyunit": "unit",
    "unit": "unit",
    "rate": "rate",
    "unitrate": "rate",
    "price": "rate",
    "discount": "discount_pct",
    "discountpercent": "discount_pct",
    "discountamount": "discount_amt",
    "total": "after_discount",
    "totalafterdiscount": "after_discount",
    "supplier": "Supplier",
    "suppliername": "Supplier",
    "vendor": "Supplier",
    "vendorname": "Supplier",
    "ponumber": "PO Number",
    "pono": "PO Number",
    "podate": "PO Date",
    "date": "PO Date",
    "gstno": "Supplier GST",
    "suppliergst": "Supplier GST",
    "panno": "Supplier PAN",
    "supplierpan": "Supplier PAN",
    "email": "Supplier Email",
    "emailid": "Supplier Email",
    "supplieremail": "Supplier Email",
    "contactno": "Supplier Contact",
    "contact": "Supplier Contact",
    "suppliercontact": "Supplier Contact",
    "paymentterm": "Payment Terms",
    "paymentterms": "Payment Terms",
    "deliverymode": "Delivery Mode",
    "deliveryschedule": "Delivery Schedule",
    "filename": "File Name",
    "cgst": "CGST",
    "sgst": "SGST",
    "igst": "IGST",
    "grandtotal": "Grand Total",
    # New template mappings
    "custname": "Supplier",
    "custcode": "CustCode",
    "custtype": "CustType",
    "addline1": "Add_Line1",
    "addline2": "Add_Line2",
    "addline3": "Add_Line3",
    "city": "City",
    "statename": "StateName",
    "countrycode": "CountryCode",
    "pincode": "PinCode",
}


def map_to_template_columns(template_headers, extracted):
    result = {}

    for header in template_headers:
        normalized = normalize_header(header)
        source_key = COLUMN_ALIASES.get(normalized)

        if source_key is None:
            # Flexible contains-based matching.
            for alias, key in COLUMN_ALIASES.items():
                if alias in normalized or normalized in alias:
                    source_key = key
                    break

        value = ""

        if source_key:
            if source_key in extracted:
                value = extracted[source_key]

        result[header] = value

    return result


# ============================================================
# MAIN
# ============================================================

def main():
    if not INPUT_EXCEL.exists():
        raise FileNotFoundError(
            f"Excel template not found: {INPUT_EXCEL}\n"
            f"Rename your Excel file to 'template.xlsx' and place it beside this script."
        )

    if not PDF_FOLDER.exists():
        raise FileNotFoundError(
            f"PDF folder not found: {PDF_FOLDER}\n"
            f"Create a folder named 'pdfs' and put all PO PDFs inside it."
        )

    pdf_files = sorted(PDF_FOLDER.glob("*.pdf"))

    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in: {PDF_FOLDER}")

    # Read Sheet1 exactly as supplied.
    wb = load_workbook(INPUT_EXCEL)
    if SOURCE_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{SOURCE_SHEET}' was not found. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws_source = wb[SOURCE_SHEET]

    headers = [
        ws_source.cell(row=1, column=c).value
        for c in range(1, ws_source.max_column + 1)
    ]

    headers = [h if h is not None else "" for h in headers]

    output_rows = []

    for pdf_path in pdf_files:
        print(f"Processing: {pdf_path.name}")

        text = get_pdf_text(pdf_path)
        # Strip terms and conditions boilerplate
        text = re.split(r"TERMS\s+(?:AND|&)\s+CONDITIONS", text, flags=re.I)[0]
        header_data = extract_header(text, pdf_path.name)
        po_num = str(header_data.get("PO Number", "")).strip()
        if po_num.lower().startswith("pulj"):
            print(f"Skipping {pdf_path.name}: PO Number '{po_num}' starts with PULJ")
            continue

        totals = extract_totals(text)
        items = extract_items(text)

        # If table extraction fails, still create one row with PO-level data.
        if not items:
            items = [{
                "sr": "",
                "qty": "",
                "unit": "",
                "rate": "",
                "rate_unit": "",
                "value": "",
                "discount_amt": "",
                "discount_pct": "",
                "after_discount": "",
                "description": ""
            }]

        for item in items:
            extracted = {
                **header_data,
                **totals,
                "sr": item.get("sr", ""),
                "qty": number(item.get("qty")),
                "unit": item.get("unit", ""),
                "rate": number(item.get("rate")),
                "rate_unit": item.get("rate_unit", ""),
                "value": number(item.get("value")),
                "discount_amt": number(item.get("discount_amt")),
                "discount_pct": number(item.get("discount_pct")),
                "after_discount": number(item.get("after_discount")),
                "description": item.get("description", ""),
            }

            output_rows.append(
                map_to_template_columns(headers, extracted)
            )

    # De-duplicate rows by CustName (Supplier Name)
    unique_rows = {}
    for row in output_rows:
        supplier_val = None
        for k, v in row.items():
            if normalize_header(k) in ["custname", "supplier", "suppliername"]:
                supplier_val = v
                break
        if supplier_val:
            key = clean(supplier_val).lower()
            if key not in unique_rows:
                unique_rows[key] = row
            else:
                existing = unique_rows[key]
                # Merge populated fields
                for k, v in row.items():
                    if not existing.get(k) and v:
                        existing[k] = v
        else:
            unique_rows[id(row)] = row
    output_rows = list(unique_rows.values())

    # Remove old output sheet if it exists.
    if OUTPUT_SHEET in wb.sheetnames:
        del wb[OUTPUT_SHEET]

    ws_out = wb.create_sheet(OUTPUT_SHEET)

    # Copy header row style from Sheet1.
    for col_idx, header in enumerate(headers, start=1):
        source_cell = ws_source.cell(row=1, column=col_idx)
        target_cell = ws_out.cell(row=1, column=col_idx, value=header)

        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)

        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format

    # Write data.
    for row_idx, row_data in enumerate(output_rows, start=2):
        for col_idx, header in enumerate(headers, start=1):
            ws_out.cell(
                row=row_idx,
                column=col_idx,
                value=row_data.get(header, "")
            )

    # Copy column widths from Sheet1.
    for col_letter, dimension in ws_source.column_dimensions.items():
        ws_out.column_dimensions[col_letter].width = dimension.width

    ws_out.freeze_panes = "A2"
    ws_out.auto_filter.ref = ws_out.dimensions

    try:
        wb.save(OUTPUT_EXCEL)
        print(f"Output file         : {OUTPUT_EXCEL}")
    except PermissionError:
        import time
        alt_output = OUTPUT_EXCEL.parent / f"{OUTPUT_EXCEL.stem}_{int(time.time())}.xlsx"
        print(f"\nWARNING: Could not save to {OUTPUT_EXCEL} because the file is open/locked in Excel.")
        print(f"Saving to alternative file: {alt_output}")
        print("Please close the locked Excel file and run the script again to overwrite it.")
        wb.save(alt_output)

    print("\nCompleted successfully.")
    print(f"PDF files processed : {len(pdf_files)}")
    print(f"Unique Suppliers    : {len(output_rows)}")


if __name__ == "__main__":
    main()
