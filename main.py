import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Step 1: Access Google Sheet
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('path/to/credentials.json', scope)
client = gspread.authorize(creds)

sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1grSGOpN0l3mCPmnDBFhGrG69BJwL8vZ5z3dg1_fOcbo/edit?pli=1&gid=1202466527#gid=1202466527')
worksheet = sheet.get_worksheet(0)

package_data = worksheet.col_values(2)

# Step 2: Capture Scanned Barcodes
scanned_barcodes = []

while True:
    barcode = input("Scan a barcode (or type 'done' to finish): ")
    if barcode.lower() == 'done':
        break
    scanned_barcodes.append(barcode)

# Step 3: Compare Scanned Barcodes with Google Sheet Data
missing_barcodes = [barcode for barcode in scanned_barcodes if barcode not in package_data]

# Step 4: Generate a Report
report = "\n".join(missing_barcodes)
print("Copy and paste the following list of missing barcodes:\n", report)
