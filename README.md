# spp-automation
scripts created to automated accounting and administrative tasks of a simple office environment

#######################################
### LEDGER UPDATE SCRIPT
## Extracts invoice data from PDFs generated using Tally Invoicing software and updates an Excel ledger

## Author: Umang Maheshwari

## Description:
This script reads PDF invoices from a specified folder, extracts relevant invoice details,
and updates an Excel ledger with the extracted information.


# This script is intended to automate the process of updating the ledger with invoice data.

# Usage:
1. Update the EXCEL_PATH variable with the absolute path to your Excel ledger file.
2. Ensure the folder containing PDF invoices is specified in cell C3 of the '001_Summary' sheet.
3. Run this script to automatically extract invoice data and update the ledger.

# Requirements
openpyxl
pypdf
#######################################



