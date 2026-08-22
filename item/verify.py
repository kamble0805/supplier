import os
import re
import sys
import logging
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

PO_FOLDER = r"c:\Users\Shubham\OneDrive\Desktop\sale register\purchase\po's\pdfs"
ITEM_MASTER_FILE = r"c:\Users\Shubham\OneDrive\Desktop\sale register\purchase\po's\item\item master.csv"
OUTPUT_FILE = r"c:\Users\Shubham\OneDrive\Desktop\sale register\purchase\po's\item\ItemMaster_Reconciled.xlsx"

ITEM_MASTER_SHEET = 0
MISSING_SHEET_NAME = "Missing PO Items"
SOURCE_SHEET_NAME = "PO Item Sources"
REPORT_SHEET_NAME = "Reconciliation Report"
COMPLETE_PO_SHEET_NAME = "Complete PO Details"


# ------------------------------------------------------------
# Change these if your column names are different
# ------------------------------------------------------------

ITEM_MASTER_ITEM_COLUMNS = [
    "ItemCode",
    "ItemNo",
    "Item Number",
    "Item Code",
    "Item No",
    "Part Number",
    "Part No",
    "Item",
]

PO_ITEM_COLUMNS = [
    "Item Number",
    "Item Code",
    "Item No",
    "Part Number",
    "Part No",
    "Material Code",
    "Material",
]


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    filename="po_reconciliation.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logger = logging.getLogger(__name__)


# ============================================================
# NORMALIZE ITEM CODE
# ============================================================

def normalize_item(value):
    """
    Normalize item numbers for comparison.

    Example:
        ABC-001 -> ABC001
        abc 001 -> ABC001
        ABC_001 -> ABC001
    """

    if pd.isna(value):
        return ""

    value = str(value).strip().upper()

    # Remove Excel-style .0
    if value.endswith(".0"):
        value = value[:-2]

    # Remove spaces and separators
    value = re.sub(r"[\s\-_\/\\\.]+", "", value)

    # Keep only letters and numbers
    value = re.sub(r"[^A-Z0-9]", "", value)

    return value


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(columns, possible_names):

    normalized_columns = {
        str(col).strip().lower(): col
        for col in columns
    }

    for name in possible_names:
        key = name.strip().lower()

        if key in normalized_columns:
            return normalized_columns[key]

    # Partial matching
    for col in columns:
        col_clean = str(col).strip().lower()

        for name in possible_names:
            if name.strip().lower() in col_clean:
                return col

    return None


# ============================================================
# LOAD ITEM MASTER
# ============================================================

def load_item_master():

    logger.info("Loading Item Master")

    if ITEM_MASTER_FILE.lower().endswith(".csv"):
        try:
            df = pd.read_csv(ITEM_MASTER_FILE, dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(ITEM_MASTER_FILE, dtype=str, encoding='latin1')
    else:
        excel = pd.ExcelFile(ITEM_MASTER_FILE)
        sheet_name = excel.sheet_names[ITEM_MASTER_SHEET]
        df = pd.read_excel(
            ITEM_MASTER_FILE,
            sheet_name=sheet_name,
            dtype=str
        )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    item_column = find_column(
        df.columns,
        ITEM_MASTER_ITEM_COLUMNS
    )

    if not item_column:
        raise ValueError(
            "Could not find Item Number column in Item Master.\n"
            f"Available columns: {list(df.columns)}"
        )

    logger.info(
        f"Item Master item column detected: {item_column}"
    )

    return df, item_column


# ============================================================
# READ PDF
# ============================================================

def extract_pdf(file_path):

    records = []

    try:

        import fitz

        document = fitz.open(file_path)

        full_text = ""

        for page in document:
            text = page.get_text("text")

            if text:
                full_text += "\n" + text

        document.close()

        if not full_text.strip():
            logger.warning(
                f"No text found in PDF: {file_path}"
            )

            return []

        return parse_text_for_items(
            full_text,
            file_path
        )

    except Exception as e:

        logger.error(
            f"Failed PDF {file_path}: {e}"
        )

        return []


# ============================================================
# READ EXCEL PO
# ============================================================

def extract_excel(file_path):

    records = []

    try:

        sheets = pd.read_excel(
            file_path,
            sheet_name=None,
            dtype=str
        )

        for sheet_name, df in sheets.items():

            if df.empty:
                continue

            df.columns = [
                str(col).strip()
                for col in df.columns
            ]

            item_column = find_column(
                df.columns,
                PO_ITEM_COLUMNS
            )

            if not item_column:
                continue

            for _, row in df.iterrows():

                item = row.get(item_column)

                if pd.isna(item):
                    continue

                item = str(item).strip()

                if not item:
                    continue

                record = {
                    "Item Number": item,
                    "PO File": Path(file_path).name,
                    "PO Number": extract_po_number(
                        str(file_path),
                        df
                    )
                }

                for col in df.columns:
                    record[col] = row.get(col)

                records.append(record)

        return records

    except Exception as e:

        logger.error(
            f"Failed Excel PO {file_path}: {e}"
        )

        return []


# ============================================================
# EXTRACT PO NUMBER
# ============================================================

def extract_po_number(file_path, df=None):

    filename = Path(file_path).stem

    # Look for PO number inside filename
    match = re.search(
        r"(?:PO|P\.O\.|PURCHASE.?ORDER)[\s\-_:#]*([A-Z0-9\/\-_]+)",
        filename,
        re.IGNORECASE
    )

    if match:
        return match.group(1)

    # Look inside dataframe columns
    if df is not None:

        po_columns = [
            "PO Number",
            "PO No",
            "Purchase Order",
            "Purchase Order No",
            "Order No",
        ]

        column = find_column(
            df.columns,
            po_columns
        )

        if column and not df.empty:

            value = df.iloc[0].get(column)

            if pd.notna(value):
                return str(value).strip()

    return filename


# ============================================================
# VALIDATE ITEM CODE
# ============================================================

def is_valid_item_code(code, line=None, po_number=None, filename=None):
    if not code:
        return False
        
    code_clean = str(code).strip().upper()
    
    # 1. Skip if it is a date
    if re.search(r"\b\d{2}[-\/\.]\d{2}[-\/\.]\d{4}\b", code_clean):
        return False
    if re.search(r"\b\d{4}[-\/\.]\d{2}[-\/\.]\d{2}\b", code_clean):
        return False
    if re.search(r"\b\d{1,2}[-\/\.][A-Z]{3}[-\/\.]\d{2,4}\b", code_clean):
        return False
        
    # 2. Skip if it's a standard PAN number
    if re.match(r"^[A-Z]{5}\d{4}[A-Z]$", code_clean):
        return False
        
    # 3. Skip if it's a GSTIN
    if re.match(r"^\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]$", code_clean):
        return False
        
    # 4. Skip if it contains a 10-digit number (highly likely a phone number)
    if re.search(r"\d{10}", code_clean):
        return False
        
    # 5. Skip if it's a PIN code starting with 40-44 (Maharashtra PINs) or any 6-digit PIN in address lines
    if re.search(r"\b4[0-4]\d{4}\b", code_clean):
        return False
        
    # 6. Skip if it looks like a financial year or PO number containing FY references
    if re.search(r"\b\d{2}-\d{2}\b", code_clean) and any(yr in code_clean for yr in ["24-25", "25-26", "26-27", "27-28"]):
        return False
    if any(yr in code_clean for yr in ["2024", "2025", "2026", "2027"]):
        if "/" in code_clean or "-" in code_clean:
            return False

    # 7. Skip if it is the PO number or filename itself
    if po_number:
        po_norm = re.sub(r"[^A-Z0-9]", "", str(po_number).upper())
        code_norm = re.sub(r"[^A-Z0-9]", "", code_clean)
        if code_norm == po_norm or code_norm in po_norm or po_norm in code_norm:
            return False
            
    if filename:
        fn_stem = Path(filename).name.upper()
        fn_norm = re.sub(r"[^A-Z0-9]", "", fn_stem)
        code_norm = re.sub(r"[^A-Z0-9]", "", code_clean)
        if code_norm == fn_norm or code_norm in fn_norm or fn_norm in code_norm:
            return False

    # 8. Skip if it's in a line that is clearly not an item line
    if line:
        line_clean = line.strip().lower()
        if "@" in line_clean or "www." in line_clean or "http" in line_clean:
            return False
        if line_clean.startswith("(") or "inr" in line_clean or "rs." in line_clean or "rs " in line_clean or "/1000" in line_clean or "/100" in line_clean:
            return False
        # Add Quotation/Serial/Reference/Address keywords to skip lines
        skip_line_words = [
            "pan no", "gst no", "state code", "requisition", "purchase order", "ref no",
            "contact no", "phone", "email", "date :", "date:", "date  :", "page",
            "quotation", "quot no", "quot. no", "sr. no", "sr.no", "serial no", "serial number",
            "sl. no", "sl.no", "sl no", "site v", "surajpur", "kasna", "ambernath", "midc", "nagar"
        ]
        if any(word in line_clean for word in skip_line_words):
            return False
            
    return True


# ============================================================
# PARSE PDF TEXT
# ============================================================

def parse_text_for_items(text, file_path):

    records = []

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    # Possible item-code patterns
    item_patterns = [
        r"\b[A-Z]{1,10}[-_/]?\d{2,}[A-Z0-9\-_/]*\b",
        r"\b\d{2,}[-_/][A-Z0-9\-_/]+\b",
    ]

    po_number = extract_po_number(file_path)

    for line in lines:

        # Skip obvious non-item lines
        skip_words = [
            "purchase order",
            "supplier",
            "vendor",
            "gst",
            "tax",
            "subtotal",
            "total",
            "amount",
            "freight",
            "discount",
            "grand total",
            "terms",
            "bank",
            "address",
            "road",
            "estate",
            "dist",
            "maharashtra",
            "india",
            "shop",
            "premises",
            "opp",
            "floor",
            "building",
            "plot",
            "sector",
            "ward",
            "lane",
            "street",
            "phone",
            "tel",
            "email",
            "contact",
            "requisition",
            "ref no",
            "hsn",
            "sac",
            "delivery",
            "payment",
            "signatory",
            "authorised",
            "note",
            "packing",
            "forwarding",
            "date :",
            "date:",
            "pan no",
            "pan:",
            "pan  :",
            "certified company",
            "state code",
            "company state",
            "compound",
            "behind",
            "near",
            "opposite",
            "beside",
            "next to",
            "industrial estate",
            "chambers",
            "plaza",
            "complex",
            "highway",
            "inr",
            "rs.",
            "rs ",
            "/1000",
            "/100",
            "industrial area",
            "site v",
            "surajpur",
            "kasna",
            "midc",
            "nagar",
            "ambernath",
            "bhiwandi",
            "thane",
            "mumbai",
            "palghar",
            "vasai",
            "bhayander",
            "dist.",
            "taluka",
            "tal.",
            "quot",
            "quotation",
            "length",
            "thickness",
            "width",
            "height",
            "od-",
            "id-",
            "micron",
            "microns",
            "size-"
        ]

        if any(
            word in line.lower()
            for word in skip_words
        ):
            continue

        item_found = None

        for pattern in item_patterns:

            match = re.search(
                pattern,
                line,
                re.IGNORECASE
            )

            if match:
                item_found = match.group(0)
                break

        if not item_found or not is_valid_item_code(item_found, line=line, po_number=po_number, filename=file_path):
            continue

        records.append({
            "Item Number": item_found,
            "Description": line,
            "PO File": Path(file_path).name,
            "PO Number": po_number,
        })

    return records


# ============================================================
# PROCESS PO FOLDER
# ============================================================

def process_po_folder():

    all_records = []

    po_files = []

    for extension in [
        "*.pdf",
        "*.xlsx",
        "*.xls",
        "*.csv"
    ]:

        po_files.extend(
            Path(PO_FOLDER).glob(extension)
        )

    po_files = sorted(
        set(po_files)
    )

    logger.info(
        f"Found {len(po_files)} PO files"
    )

    failed_files = []

    for index, file_path in enumerate(
        po_files,
        start=1
    ):

        print(
            f"Processing {index}/{len(po_files)}: "
            f"{file_path.name}"
        )

        try:

            extension = file_path.suffix.lower()

            if extension == ".pdf":

                records = extract_pdf(
                    file_path
                )

            elif extension in [
                ".xlsx",
                ".xls"
            ]:

                records = extract_excel(
                    file_path
                )

            elif extension == ".csv":

                df = pd.read_csv(
                    file_path,
                    dtype=str
                )

                item_column = find_column(
                    df.columns,
                    PO_ITEM_COLUMNS
                )

                records = []

                if item_column:

                    for _, row in df.iterrows():

                        item = row.get(
                            item_column
                        )

                        if pd.isna(item):
                            continue

                        records.append({
                            "Item Number": str(
                                item
                            ).strip(),
                            "PO File": file_path.name,
                            "PO Number": extract_po_number(
                                str(file_path),
                                df
                            )
                        })

            else:

                continue

            if records:
                all_records.extend(records)

            else:
                logger.warning(
                    f"No item records found: "
                    f"{file_path.name}"
                )

        except Exception as e:

            logger.exception(
                f"Failed: {file_path.name}"
            )

            failed_files.append({
                "PO File": file_path.name,
                "Error": str(e)
            })

    return all_records, failed_files


# ============================================================
# CREATE MISSING ITEMS
# ============================================================

def identify_missing_items(
    item_master,
    item_master_item_column,
    po_records
):

    print("\nComparing PO items with Item Master...")

    # Existing Item Master items
    master_items = set()

    # Match against both ItemNo and ItemCode if they exist, to avoid false positives on missing items
    columns_to_check = [c for c in ["ItemNo", "ItemCode", item_master_item_column] if c in item_master.columns]
    for col in columns_to_check:
        for value in item_master[col]:
            normalized = normalize_item(value)
            if normalized:
                master_items.add(normalized)

    missing_records = []
    source_records = []

    seen_missing = set()

    for record in po_records:

        original_item = record.get(
            "Item Number",
            ""
        )

        normalized_item = normalize_item(
            original_item
        )

        if not normalized_item:
            continue

        # Already exists
        if normalized_item in master_items:
            continue

        # Source tracking
        source_records.append({
            "Item Number": original_item,
            "Normalized Item Number": normalized_item,
            "PO File": record.get(
                "PO File",
                ""
            ),
            "PO Number": record.get(
                "PO Number",
                ""
            )
        })

        # Avoid duplicate missing items
        if normalized_item in seen_missing:
            continue

        seen_missing.add(
            normalized_item
        )

        missing_records.append(
            record
        )

    return (
        missing_records,
        source_records
    )


# ============================================================
# MAP PO DATA TO ITEM MASTER
# ============================================================

def map_to_item_master(
    missing_records,
    item_master_columns
):

    output_records = []

    # Find the main item column from the Item Master columns list
    item_no_col = find_column(item_master_columns, ITEM_MASTER_ITEM_COLUMNS)

    for record in missing_records:

        new_record = {
            column: ""
            for column in item_master_columns
        }

        for master_column in item_master_columns:

            master_clean = (
                str(master_column)
                .strip()
                .lower()
            )

            # 1. Direct column match
            value_found = False
            for po_column, value in record.items():

                po_clean = (
                    str(po_column)
                    .strip()
                    .lower()
                )

                if master_clean == po_clean:

                    if pd.notna(value):
                        new_record[
                            master_column
                        ] = value
                        value_found = True

                    break

            if value_found:
                continue

            # 2. Match item code/number columns
            is_item_identifier = (
                master_clean in [
                    "itemno",
                    "itemcode",
                    "item number",
                    "item code",
                    "item no",
                    "part number",
                    "part no",
                    "item"
                ] or master_column == item_no_col
            )

            if is_item_identifier:
                new_record[master_column] = record.get("Item Number", "")

            # 3. Match description columns
            elif "desc" in master_clean or "description" in master_clean:
                new_record[master_column] = record.get("Description", "")

        output_records.append(
            new_record
        )

    return pd.DataFrame(
        output_records,
        columns=item_master_columns
    )


# ============================================================
# SAVE EXCEL
# ============================================================

def save_output(
    item_master,
    missing_df,
    source_records,
    po_records,
    failed_files,
    total_po_files,
    total_po_lines
):

    print("\nCreating output workbook...")

    po_df = pd.DataFrame(po_records)
    if not po_df.empty:
        first_cols = ["PO Number", "PO File", "Item Number", "Description"]
        existing_first_cols = [c for c in first_cols if c in po_df.columns]
        other_cols = [c for c in po_df.columns if c not in existing_first_cols]
        po_df = po_df[existing_first_cols + other_cols]

    # Resolve permission error if output file is open/locked
    target_output_file = OUTPUT_FILE
    output_path = Path(OUTPUT_FILE)
    for i in range(100):
        try:
            if os.path.exists(target_output_file):
                with open(target_output_file, 'a'):
                    pass
            break
        except PermissionError:
            target_output_file = str(output_path.parent / f"{output_path.stem}_{i+1}{output_path.suffix}")

    if target_output_file != OUTPUT_FILE:
        print(f"\nWARNING: '{OUTPUT_FILE}' is locked or open in Excel.")
        print(f"Saving output to alternative file: '{target_output_file}'")

    if ITEM_MASTER_FILE.lower().endswith(".csv"):
        # Since CSV is a plain text file, we cannot load/edit it using openpyxl.
        # Instead, we write a fresh Excel workbook with the Item Master as the first sheet
        # and then append/write the reconciliation sheets.
        with pd.ExcelWriter(
            target_output_file,
            engine="openpyxl",
            mode="w"
        ) as writer:
            item_master.to_excel(
                writer,
                sheet_name="Item Master",
                index=False
            )
            
            po_df.to_excel(
                writer,
                sheet_name=COMPLETE_PO_SHEET_NAME,
                index=False
            )

            missing_df.to_excel(
                writer,
                sheet_name=MISSING_SHEET_NAME,
                index=False
            )

            source_df = pd.DataFrame(
                source_records
            )

            source_df.to_excel(
                writer,
                sheet_name=SOURCE_SHEET_NAME,
                index=False
            )

            report_data = [
                ["Total PO Files", total_po_files],
                ["PO Files Failed", len(failed_files)],
                ["Total PO Item Lines", total_po_lines],
                ["Unique Missing Items", len(missing_df)],
                ["PO Items Found in Item Master", total_po_lines - len(source_records)],
                ["Missing Item Occurrences", len(source_records)],
                ["Unique Missing Items", len(missing_df)],
            ]

            report_df = pd.DataFrame(
                report_data,
                columns=["Metric", "Count"]
            )

            report_df.to_excel(
                writer,
                sheet_name=REPORT_SHEET_NAME,
                index=False
            )

            if failed_files:
                failed_df = pd.DataFrame(
                    failed_files
                )
                start_row = len(report_df) + 4
                failed_df.to_excel(
                    writer,
                    sheet_name=REPORT_SHEET_NAME,
                    index=False,
                    startrow=start_row
                )
    else:
        # Load original workbook
        workbook = load_workbook(
            ITEM_MASTER_FILE
        )

        # Remove old generated sheets
        for sheet_name in [
            MISSING_SHEET_NAME,
            SOURCE_SHEET_NAME,
            REPORT_SHEET_NAME,
            COMPLETE_PO_SHEET_NAME
        ]:

            if sheet_name in workbook.sheetnames:

                del workbook[
                    sheet_name
                ]

        # Save first so pandas can append sheets
        workbook.save(
            target_output_file
        )

        # Write sheets
        with pd.ExcelWriter(
            target_output_file,
            engine="openpyxl",
            mode="a"
        ) as writer:

            po_df.to_excel(
                writer,
                sheet_name=COMPLETE_PO_SHEET_NAME,
                index=False
            )

            missing_df.to_excel(
                writer,
                sheet_name=MISSING_SHEET_NAME,
                index=False
            )

            source_df = pd.DataFrame(
                source_records
            )

            source_df.to_excel(
                writer,
                sheet_name=SOURCE_SHEET_NAME,
                index=False
            )

            report_data = [

                [
                    "Total PO Files",
                    total_po_files
                ],

                [
                    "PO Files Failed",
                    len(failed_files)
                ],

                [
                    "Total PO Item Lines",
                    total_po_lines
                ],

                [
                    "Unique Missing Items",
                    len(missing_df)
                ],

                [
                    "PO Items Found in Item Master",
                    total_po_lines -
                    len(source_records)
                ],

                [
                    "Missing Item Occurrences",
                    len(source_records)
                ],

                [
                    "Unique Missing Items",
                    len(missing_df)
                ],
            ]

            report_df = pd.DataFrame(
                report_data,
                columns=[
                    "Metric",
                    "Count"
                ]
            )

            report_df.to_excel(
                writer,
                sheet_name=REPORT_SHEET_NAME,
                index=False
            )

            # Failed files
            if failed_files:

                failed_df = pd.DataFrame(
                    failed_files
                )

                start_row = len(report_df) + 4

                failed_df.to_excel(
                    writer,
                    sheet_name=REPORT_SHEET_NAME,
                    index=False,
                    startrow=start_row
                )

    print(
        f"\nOutput created:\n{target_output_file}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 60)
    print("PO -> ITEM MASTER RECONCILIATION")
    print("=" * 60)

    # Validate paths
    if not os.path.exists(
        PO_FOLDER
    ):

        raise FileNotFoundError(
            f"PO folder not found:\n{PO_FOLDER}"
        )

    if not os.path.exists(
        ITEM_MASTER_FILE
    ):

        raise FileNotFoundError(
            f"Item Master not found:\n"
            f"{ITEM_MASTER_FILE}"
        )

    # Load Item Master
    (
        item_master,
        item_master_item_column
    ) = load_item_master()

    print(
        f"Item Master records: "
        f"{len(item_master)}"
    )

    print(
        f"Item column: "
        f"{item_master_item_column}"
    )

    # Process POs
    (
        po_records,
        failed_files
    ) = process_po_folder()

    print(
        f"\nPO item lines extracted: "
        f"{len(po_records)}"
    )

    # Compare
    (
        missing_records,
        source_records
    ) = identify_missing_items(
        item_master,
        item_master_item_column,
        po_records
    )

    # Convert to Item Master format
    missing_df = map_to_item_master(
        missing_records,
        list(item_master.columns)
    )

    print(
        f"Unique missing items: "
        f"{len(missing_df)}"
    )

    # Save
    save_output(
        item_master=item_master,
        missing_df=missing_df,
        source_records=source_records,
        po_records=po_records,
        failed_files=failed_files,
        total_po_files=len(
            list(
                Path(PO_FOLDER).glob("*")
            )
        ),
        total_po_lines=len(
            po_records
        )
    )

    print("\nPROCESS COMPLETED")
    print("=" * 60)


if __name__ == "__main__":
    main()