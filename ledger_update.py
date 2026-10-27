#######################################
# LEDGER UPDATE SCRIPT
# Extracts invoice data from PDFs generated using Tally Invoicing software and updates an Excel ledger

# Author: Umang Maheshwari
# Date: 2026-10-26

# Description:
# This script reads PDF invoices from a specified folder, extracts relevant invoice details,
# and updates an Excel ledger with the extracted information.


#  This script is intended to automate the process of updating the ledger with invoice data.

# Usage:
# 1. Update the EXCEL_PATH variable with the absolute path to your Excel ledger file.
# 2. Ensure the folder containing PDF invoices is specified in cell C3 of the '001_Summary' sheet.
# 3. Run this script to automatically extract invoice data and update the ledger.
#######################################

from datetime import datetime
import os
import re
import openpyxl
from openpyxl.worksheet.table import Table
from pypdf import PdfReader

# ==========================================
# CONFIGURATION
# ==========================================
EXCEL_PATH = r"path/to/your/excel/file.xlsx"  # Update with your absolute path


def extract_invoice_data(pdf_path):
    """Extracts Buyer Name, GSTIN, Invoice No, Dated, and Total Amount from PDF invoice."""
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return None

    data = {
        "buyer_name": "Unknown Buyer",
        "gstin": "",
        "invoice_no": "",
        "dated": "",
        "total_amount": 0.0,
    }

    # 1. Extract Buyer Name (typically after "Buyer (Bill to)")
    buyer_match = re.search(
        r"Buyer\s*\(Bill\s*to\)\s*\n([^\n]+)(?:\n([^\n]+))?", text, re.IGNORECASE
    )
    if buyer_match:
        lines = [
            l.strip()
            for l in buyer_match.groups()
            if l and l.strip() and "|" not in l
        ]
        if lines:
            data["buyer_name"] = lines[0]

    # 2. Extract GSTIN/UIN (Buyer GSTIN)
    all_gstins = re.findall(
        r"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}", text
    )
    if len(all_gstins) > 1:
        data["gstin"] = all_gstins[1]
    elif len(all_gstins) == 1:
        data["gstin"] = all_gstins[0]

    # 3. Extract Invoice No.
    inv_match = re.search(
        r"Invoice\s*No\.?\s*\n.*?([A-Za-z0-9\-\/]+)", text, re.IGNORECASE
    ) or re.search(
        r"Invoice\s*No\.?\s*[:|]?\s*([A-Za-z0-9\-\/]+)", text, re.IGNORECASE
    )
    if inv_match:
        data["invoice_no"] = inv_match.group(1).strip()

    # 4. Extract Dated
    date_match = re.search(
        r"Dated\s*\n.*?([0-9]{1,2}-[A-Za-z]{3}-[0-9]{2,4})", text, re.IGNORECASE
    ) or re.search(
        r"Dated\s*[:|]?\s*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{2,4})", text, re.IGNORECASE
    )
    if date_match:
        data["dated"] = date_match.group(1).strip()

    # 5. Extract Total Amount following ₹ or Rs. symbol
    amt_match = re.search(r"(?:₹|Rs\.?)\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if amt_match:
        try:
            amt_str = amt_match.group(1).replace(",", "")
            data["total_amount"] = float(amt_str)
        except ValueError:
            pass

    return data


def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"Error: Excel file not found at {EXCEL_PATH}")
        return

    wb = openpyxl.load_workbook(EXCEL_PATH)

    # 1. Check sheet 001_Summary for folder path in cell C3
    if "001_Summary" not in wb.sheetnames:
        print("Error: Sheet '001_Summary' not found.")
        return

    ws_summary = wb["001_Summary"]
    folder_path = ws_summary["C3"].value

    if not folder_path or not os.path.exists(str(folder_path)):
        print(
            f"Error: Invalid or missing folder path in '001_Summary' cell C3: {folder_path}"
        )
        return

    folder_path = str(folder_path)

    # 2. Get list of PDF invoices in folder
    invoice_files = [
        f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")
    ]

    # 3. Setup Processed Sheet (000_Processed_invoices)
    if "000_Processed_invoices" not in wb.sheetnames:
        ws_proc = wb.create_sheet("000_Processed_invoices")
        ws_proc.append(["Date", "Invoice Name", "Creation Date", "Processed"])
    else:
        ws_proc = wb["000_Processed_invoices"]

    processed_dict = {}
    for row in range(2, ws_proc.max_row + 1):
        inv_name = ws_proc.cell(row=row, column=2).value
        status = ws_proc.cell(row=row, column=4).value
        if inv_name:
            processed_dict[str(inv_name).strip()] = (
                str(status).strip().lower() == "yes"
            )

    current_date_str = datetime.now().strftime("%Y-%m-%d")
    for fname in invoice_files:
        if fname not in processed_dict:
            file_path = os.path.join(folder_path, fname)
            
            # Fetch OS file creation date (fallback to modification time if ctime is unreliable)
            try:
                ctime = os.path.getctime(file_path)
            except Exception:
                ctime = os.path.getmtime(file_path)
            file_creation_date = datetime.fromtimestamp(ctime).strftime("%Y-%m-%d")

            ws_proc.append([current_date_str, fname, file_creation_date, "No"])
            processed_dict[fname] = False

    wb.save(EXCEL_PATH)

    # 4. Process unprocessed invoices
    for row in range(2, ws_proc.max_row + 1):
        inv_name = ws_proc.cell(row=row, column=2).value
        status = ws_proc.cell(row=row, column=4).value

        if not inv_name:
            continue

        if status and str(status).strip().lower() == "yes":
            continue

        pdf_full_path = os.path.join(folder_path, str(inv_name))
        if not os.path.exists(pdf_full_path):
            print(f"Warning: File {inv_name} not found in folder. Skipping.")
            continue

        # 5. Extract data from invoice
        extracted = extract_invoice_data(pdf_full_path)
        if not extracted or not extracted["buyer_name"]:
            print(f"Could not extract details for {inv_name}")
            continue

        client_name = extracted["buyer_name"].strip()
        client_gst = extracted["gstin"].strip()
        invoice_no = extracted["invoice_no"].strip()
        invoice_date = extracted["dated"].strip()
        invoice_amt = extracted["total_amount"]

        # 6. Check if client exists in 001_Summary table
        client_row_in_summary = None
        existing_clients_count = 0

        for r in range(7, ws_summary.max_row + 1):
            c_name = ws_summary.cell(row=r, column=2).value
            if c_name:
                existing_clients_count += 1
                if str(c_name).strip().lower() == client_name.lower():
                    client_row_in_summary = r
                    break

        # 7. If client does not exist, copy 999_Template and create client sheet with tables & updated formulas
        if not client_row_in_summary:
            if "999_Template" not in wb.sheetnames:
                print("Error: '999_Template' sheet is missing.")
                return

            ws_template = wb["999_Template"]

            existing_numbers = []
            for sheetname in wb.sheetnames:
                match = re.match(r"^(\d{3})_", sheetname)
                if match:
                    existing_numbers.append(int(match.group(1)))

            next_num = 1
            while next_num in existing_numbers or next_num == 999:
                next_num += 1

            clean_client_tag = re.sub(r"[^\w\s]", "", client_name)
            clean_client_tag = re.sub(
                r"\s+", "_", clean_client_tag.strip()
            ).lower()
            new_sheet_name = f"{next_num:03d}_{clean_client_tag}"[:31]

            new_ws = wb.copy_worksheet(ws_template)
            new_ws.title = new_sheet_name

            # Copy tables and update formulas in cells C4 and I4
            template_tables = (
                ws_template._tables.values()
                if hasattr(ws_template, "_tables")
                else ws_template.tables.values()
            )
            for tbl in template_tables:
                if hasattr(tbl, "ref"):
                    tbl_name = getattr(tbl, "name", "Table")
                    new_table_name = f"{tbl_name}_{next_num}"
                    new_tbl = Table(displayName=new_table_name, ref=tbl.ref)
                    if tbl.tableStyleInfo:
                        new_tbl.tableStyleInfo = tbl.tableStyleInfo
                    new_ws.add_table(new_tbl)

                    if "Table3" in tbl_name:
                        new_ws["C4"] = f"=SUM({new_table_name}[Amount])"
                    elif "Table4" in tbl_name:
                        new_ws["I4"] = f"=SUM({new_table_name}[Amount])"

            # Set Client Name in C1 and GST in C2
            new_ws["C1"] = client_name
            new_ws["C2"] = client_gst

            next_summary_row = 7 + existing_clients_count
            ws_summary.cell(row=next_summary_row, column=1, value=next_num)

            cell_client = ws_summary.cell(
                row=next_summary_row, column=2, value=client_name
            )
            # Hyperlink pointing to cell C1 of the client sheet
            cell_client.hyperlink = f"#'{new_sheet_name}'!C1"
            cell_client.style = "Hyperlink"

            ws_summary.cell(row=next_summary_row, column=3, value=client_gst)
            # Total Billing Value referencing the cell C4 of the client sheet (which sums the client's table)
            ws_summary.cell(
                row=next_summary_row,
                column=4,
                value=f"='{new_sheet_name}'!C4",
            )
            ws_summary.cell(
                row=next_summary_row,
                column=5,
                value=f"='{new_sheet_name}'!I4")

            client_row_in_summary = next_summary_row
            target_ws = new_ws
        else:
            client_name_val = ws_summary.cell(
                row=client_row_in_summary, column=2
            ).value
            target_ws = None
            for sname in wb.sheetnames:
                if client_name_val and client_name_val.lower() in sname.lower():
                    target_ws = wb[sname]
                    break
            if not target_ws:
                idx_val = ws_summary.cell(
                    row=client_row_in_summary, column=1
                ).value
                for sname in wb.sheetnames:
                    if sname.startswith(f"{int(idx_val):03d}_"):
                        target_ws = wb[sname]
                        break

        if not target_ws:
            print(f"Error: Could not locate sheet for client {client_name}")
            continue

        # 9. Enter invoice details into client sheet table
        target_row = 7
        while target_ws.cell(row=target_row, column=2).value is not None:
            target_row += 1

        sr_no = target_row - 6

        target_ws.cell(row=target_row, column=1, value=sr_no)
        target_ws.cell(row=target_row, column=2, value=invoice_no)
        target_ws.cell(row=target_row, column=3, value=invoice_date)
        target_ws.cell(row=target_row, column=4, value=invoice_amt)

        target_tables = (
            target_ws._tables.values()
            if hasattr(target_ws, "_tables")
            else target_ws.tables.values()
        )
        for tbl in target_tables:
            if hasattr(tbl, "name") and tbl.name.startswith("Table3"):
                start_cell = tbl.ref.split(":")[0]
                end_col = re.findall(r"[A-Z]+", tbl.ref)[-1]
                tbl.ref = f"{start_cell}:{end_col}{target_row}"

        # 11. Mark invoice as processed
        ws_proc.cell(row=row, column=4, value="Yes")
        wb.save(EXCEL_PATH)

    print("operation complete")


if __name__ == "__main__":
    main()