"""
Script สำหรับเตรียม Google Sheets จากไฟล์ Excel
รัน 1 ครั้งตอน setup เท่านั้น

Usage:
    python prepare_gsheet.py --excel "../Material_Status_Summary_Improved.xlsx"
"""
import argparse
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json, sys

SHEET_NAME = "Material_Status_Dashboard"

def main(excel_path, creds_path):
    # Load credentials
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(creds_path, scopes=scope)
    gc = gspread.authorize(creds)

    # Read all RS sheets from Excel → flatten to single Data table
    all_rows = []
    xls = pd.ExcelFile(excel_path)

    for sheet in ["RS1-2","RS3-4","RS5-6","RS7-8"]:
        if sheet not in xls.sheet_names:
            print(f"  ⚠ Sheet {sheet} not found, skipping")
            continue
        df_raw = pd.read_excel(excel_path, sheet_name=sheet, header=None)

        # Row 2 (index 1) = Riser group names  (B=RS1-CHWF, G=RS1-CHWR, L=RS2-CHWF, ...)
        # Row 3 (index 2) = column headers
        # Data starts row 4 (index 3)
        # Groups: cols 1-5, 6-10, 11-15, 16-20  (0-indexed: B=1)
        group_starts = [1, 7, 13, 19]  # 0-indexed; each group now = 6 cols (added Dwg Status)

        for gs in group_starts:
            group_header = str(df_raw.iloc[1, gs]).strip() if gs < df_raw.shape[1] else ""
            if not group_header or group_header == "nan":
                continue
            parts = group_header.split("-")
            riser  = parts[0] if len(parts) >= 1 else ""
            system = parts[1] if len(parts) >= 2 else ""

            for row_idx in range(3, len(df_raw)):
                row = df_raw.iloc[row_idx]
                if gs >= len(row): continue
                item_no  = row.iloc[gs]
                size     = row.iloc[gs+1] if gs+1 < len(row) else ""
                material = row.iloc[gs+2] if gs+2 < len(row) else ""
                recv     = row.iloc[gs+3] if gs+3 < len(row) else ""
                work     = row.iloc[gs+4] if gs+4 < len(row) else ""
                dwg      = row.iloc[gs+5] if gs+5 < len(row) else ""

                if pd.isna(item_no) or str(item_no).strip() == "":
                    continue

                all_rows.append({
                    "Riser":            str(riser).strip(),
                    "System":           str(system).strip(),
                    "ITEM No.":         str(int(item_no)) if not pd.isna(item_no) else "",
                    "Size":             str(size).strip() if not pd.isna(size) else "",
                    "Material":         str(material).strip() if not pd.isna(material) else "",
                    "Receiving Status": "" if pd.isna(recv) else str(recv).strip(),
                    "Work Status":      "" if pd.isna(work) else str(work).strip(),
                    "Dwg Status":       "" if pd.isna(dwg)  else str(dwg).strip(),
                })

    df_out = pd.DataFrame(all_rows)
    print(f"  → {len(df_out)} rows extracted from Excel")

    # Create or open Google Sheet
    try:
        sh = gc.open(SHEET_NAME)
        print(f"  → Opened existing sheet: {SHEET_NAME}")
    except gspread.SpreadsheetNotFound:
        sh = gc.create(SHEET_NAME)
        print(f"  → Created new sheet: {SHEET_NAME}")
        # share with your Google account
        sh.share("", perm_type="anyone", role="writer")

    # Write Data worksheet
    try:
        ws = sh.worksheet("Data")
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Data", rows=1000, cols=10)

    # Write header + data
    headers = ["Riser","System","ITEM No.","Size","Material","Receiving Status","Work Status","Dwg Status"]
    data_to_write = [headers] + df_out[headers].values.tolist()
    ws.update("A1", data_to_write)
    print(f"  ✓ Written {len(df_out)} rows to '{SHEET_NAME}/Data'")
    print(f"  ✓ Sheet URL: {sh.url}")
    print(f"\n  ⚠ Remember to share this sheet with your service account email!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel",  default="../Material_Status_Summary_Improved.xlsx")
    parser.add_argument("--creds",  default="credentials.json", help="Path to service account JSON")
    args = parser.parse_args()
    main(args.excel, args.creds)
