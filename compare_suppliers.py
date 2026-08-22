import re
import difflib
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
EXTRACTED_EXCEL = BASE_DIR / "template_PDF_Extracted.xlsx"
MASTER_CSV = BASE_DIR / "supplierpmtrack.csv"
REPORT_EXCEL = BASE_DIR / "supplier_comparison_report.xlsx"

def normalize_name(name):
    if not isinstance(name, str):
        return ""
    # Lowercase, merge spaces, strip whitespace
    return re.sub(r"\s+", " ", name).strip().lower()

def main():
    if not EXTRACTED_EXCEL.exists():
        print(f"Error: Extracted PDF Excel file not found at: {EXTRACTED_EXCEL}")
        return

    if not MASTER_CSV.exists():
        print(f"Error: Supplier master CSV file not found at: {MASTER_CSV}")
        return

    print("Loading data...")
    # Load sheets
    try:
        df_ext = pd.read_excel(EXTRACTED_EXCEL, sheet_name="PDF Extracted Data")
    except PermissionError:
        import glob
        alt_files = sorted(glob.glob(str(EXTRACTED_EXCEL.parent / "template_PDF_Extracted_*.xlsx")))
        if alt_files:
            latest_alt = Path(alt_files[-1])
            print(f"WARNING: {EXTRACTED_EXCEL.name} is locked in Excel. Loading from alternative: {latest_alt.name}")
            df_ext = pd.read_excel(latest_alt, sheet_name="PDF Extracted Data")
        else:
            print(f"Error: {EXTRACTED_EXCEL.name} is locked and no alternative files were found.")
            return
    except Exception as e:
        print(f"Error reading 'PDF Extracted Data' sheet from {EXTRACTED_EXCEL}: {e}")
        return

    try:
        try:
            df_master = pd.read_csv(MASTER_CSV, encoding='utf-8-sig')
        except UnicodeDecodeError:
            df_master = pd.read_csv(MASTER_CSV, encoding='latin1')
    except Exception as e:
        print(f"Error reading {MASTER_CSV.name}: {e}")
        return

    # Extract supplier name lists
    ext_names = df_ext["CustName"].dropna().unique().tolist()
    master_names = df_master["CustName"].dropna().unique().tolist()

    print(f"Loaded {len(ext_names)} unique suppliers from PDF Extracted Data.")
    print(f"Loaded {len(master_names)} unique suppliers from {MASTER_CSV.name}.")

    # Create normalized lookup dictionary
    master_lookup = {normalize_name(name): name for name in master_names}

    results = []

    for name in ext_names:
        norm_name = normalize_name(name)
        
        # Check exact normalized match
        if norm_name in master_lookup:
            results.append({
                "Extracted Name": name,
                "Status": "Present",
                "Matched Master Name": master_lookup[norm_name]
            })
        else:
            results.append({
                "Extracted Name": name,
                "Status": "Missing",
                "Matched Master Name": ""
            })

    df_results = pd.DataFrame(results)

    # Print Summary Table
    print("\n" + "="*80)
    print("                      SUPPLIER VERIFICATION SUMMARY")
    print("="*80)
    
    missing = df_results[df_results["Status"] == "Missing"]
    present = df_results[df_results["Status"] == "Present"]

    print(f"Present in Master  : {len(present)}")
    print(f"Missing in Master  : {len(missing)}")
    print("="*80)

    if not missing.empty:
        print("\nMISSING SUPPLIERS:")
        for idx, row in missing.iterrows():
            print(f"- {row['Extracted Name']}")

    # Filter df_ext to get complete records for missing suppliers
    missing_names = df_results[df_results["Status"] == "Missing"]["Extracted Name"].tolist()
    missing_norm = {normalize_name(name) for name in missing_names}
    df_missing_records = df_ext[df_ext["CustName"].apply(lambda x: normalize_name(x) in missing_norm)]

    # Save to Excel
    try:
        with pd.ExcelWriter(REPORT_EXCEL, engine='openpyxl') as writer:
            df_results.to_excel(writer, index=False, sheet_name="Verification Report")
            df_missing_records.to_excel(writer, index=False, sheet_name="Sheet2")
        print("\n" + "="*80)
        print(f"Detailed comparison report saved to:\n{REPORT_EXCEL}")
        print("="*80)
    except PermissionError:
        print(f"\nWarning: Could not save report to {REPORT_EXCEL} because it is open in Excel.")
        import time
        alt_report = REPORT_EXCEL.parent / f"{REPORT_EXCEL.stem}_{int(time.time())}.xlsx"
        with pd.ExcelWriter(alt_report, engine='openpyxl') as writer:
            df_results.to_excel(writer, index=False, sheet_name="Verification Report")
            df_missing_records.to_excel(writer, index=False, sheet_name="Sheet2")
        print(f"Saved alternative report to: {alt_report}")

if __name__ == "__main__":
    main()
