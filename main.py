import sys
import pandas as pd
import pdfplumber
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QFileDialog, QMessageBox, QListWidget, QSplitter
)
from PyQt5.QtCore import Qt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Function to load package data from PDF and handle multiple pages
def load_package_data_from_pdf(file_path):
    try:
        all_tables = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_tables.extend(table[1:] if all_tables else table)
        if not all_tables:
            raise ValueError("No tables found in PDF")

        df = pd.DataFrame(all_tables[1:], columns=all_tables[0])
        if 'Package' not in df.columns:
            raise KeyError("'Package' column missing in PDF")
        
        df['Package'] = df['Package'].fillna('').str.strip().str.upper()
        return df
    except Exception as e:
        print(f"Error loading PDF file: {e}")
        return pd.DataFrame()

# Function to export PDF report
def export_pdf_report(matched, unmatched, missing, output_path):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    margin = 50
    line_height = 12  # Reduced line height to fit more content
    y = height - margin

    def add_new_page():
        nonlocal y
        c.showPage()
        y = height - margin
        c.setFont("Helvetica-Bold", 14)  # Apply font changes on every page
        c.drawString(margin, y, "Barcode Scan Report (continued)")
        y -= 30  # Add spacing after title for new pages

    def add_section(title, items, item_format):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)  # Font size for section title
        c.drawString(margin, y, title)
        y -= line_height
        c.setFont("Helvetica", 9)  # Font size for content
        for item in items:
            if y < margin + line_height * 3:  # Add space for new page if content is too close to bottom
                add_new_page()
            c.drawString(margin + 10, y, item_format(item))
            y -= line_height
        y -= 15  # Add extra spacing between sections

    # Title and Summary for the first page
    c.setFont("Helvetica-Bold", 14)  # Font size for title
    c.drawString(margin, y, "Barcode Scan Report")
    y -= 25
    c.setFont("Helvetica", 10)  # Font size for summary
    c.drawString(margin, y, f"Total Scanned: {len(matched) + len(unmatched)}")
    y -= line_height
    c.drawString(margin, y, f"Matched: {len(matched)}")
    y -= line_height
    c.drawString(margin, y, f"Inactive Tags Scanned: {len(unmatched)}")
    y -= line_height
    c.drawString(margin, y, f"Missing: {len(missing)}")
    y -= 20

    # Ensure enough space for sections
    add_section("Matched Barcodes:", matched, lambda x: f"{x[0]} - {x[1]} - Qty: {x[2]}")
    add_section("Inactive Tags Scanned:", unmatched, lambda x: f"{x}")
    add_section("Missing Barcodes (Not Scanned):", missing, lambda x: f"{x[0]} - {x[1]} - Qty: {x[2]}")

    c.save()

# Function to export data to an Excel file
def export_to_excel(matched, unmatched, missing, file_path):
    try:
        # Create DataFrames for each category
        matched_df = pd.DataFrame(matched, columns=["Barcode", "Item", "Quantity"]) if matched else pd.DataFrame(columns=["Barcode", "Item", "Quantity"])
        unmatched_df = pd.DataFrame(unmatched, columns=["Unmatched Barcodes"]) if unmatched else pd.DataFrame(columns=["Unmatched Barcodes"])
        missing_df = pd.DataFrame(missing, columns=["Barcode", "Item", "Quantity"]) if missing else pd.DataFrame(columns=["Barcode", "Item", "Quantity"])

        # Create a writer for Excel format
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            if not matched_df.empty:
                matched_df.to_excel(writer, sheet_name='Matched Barcodes', index=False)
            if not unmatched_df.empty:
                unmatched_df.to_excel(writer, sheet_name='Unmatched Barcodes', index=False)
            if not missing_df.empty:
                missing_df.to_excel(writer, sheet_name='Missing Barcodes', index=False)

        QMessageBox.information(None, "Success", "Excel file exported successfully!")

    except Exception as e:
        QMessageBox.critical(None, "Error", f"Failed to export Excel file: {e}")

# Main Application Class
class BarcodeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Barcode Scanner Application - Reefer Report")
        self.df_packages = pd.DataFrame()
        self.scanned_barcodes = []
        self.matched = []
        self.unmatched = []
        self.missing = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Load PDF File Section
        load_layout = QHBoxLayout()
        self.load_button = QPushButton("Load PDF File")
        self.load_button.clicked.connect(self.load_file)
        self.file_label = QLabel("No file loaded")
        load_layout.addWidget(self.load_button)
        load_layout.addWidget(self.file_label)
        layout.addLayout(load_layout)

        # Search bar for barcode lookup
        search_layout = QHBoxLayout()
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search barcode...")
        self.search_bar.textChanged.connect(self.search_bar_function)
        search_layout.addWidget(self.search_bar)
        layout.addLayout(search_layout)

        # Barcode Input Section
        barcode_layout = QHBoxLayout()
        self.barcode_input = QLineEdit()
        self.barcode_input.setPlaceholderText("Scan or enter barcode here")
        self.barcode_input.returnPressed.connect(self.scan_barcode)
        barcode_layout.addWidget(self.barcode_input)
        layout.addLayout(barcode_layout)

        # Scanned and Remaining Barcodes Lists
        splitter = QSplitter()
        self.scanned_list = QListWidget()
        self.remaining_list = QListWidget()

        # Add double-click functionality to move barcode between lists
        self.remaining_list.itemDoubleClicked.connect(self.move_item_to_scanned)
        self.scanned_list.itemDoubleClicked.connect(self.move_item_to_remaining)

        splitter.addWidget(self.remaining_list)
        splitter.addWidget(self.scanned_list)
        splitter.setSizes([200, 200])
        layout.addWidget(splitter)

        # Result Display
        self.result_display = QTextEdit()
        self.result_display.setReadOnly(True)
        layout.addWidget(self.result_display)

        # Export and Clear Buttons
        button_layout = QHBoxLayout()
        self.export_pdf_button = QPushButton("Export Report as PDF")
        self.export_pdf_button.clicked.connect(self.export_pdf)
        self.export_excel_button = QPushButton("Export Report as Excel")
        self.export_excel_button.clicked.connect(self.export_to_excel_file)
        self.clear_button = QPushButton("Clear Barcodes")
        self.clear_button.clicked.connect(self.clear_barcodes)
        button_layout.addWidget(self.export_pdf_button)
        button_layout.addWidget(self.export_excel_button)
        button_layout.addWidget(self.clear_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)")
        if file_path:
            self.file_label.setText(file_path)
            self.df_packages = load_package_data_from_pdf(file_path)
            if not self.df_packages.empty:
                self.update_remaining_list()
                QMessageBox.information(self, "Success", "PDF file loaded successfully!")
            else:
                QMessageBox.critical(self, "Error", "Failed to load PDF file or 'Package' column missing.")

    def update_remaining_list(self):
        self.remaining_list.clear()
        for package in self.df_packages['Package'].values:
            self.remaining_list.addItem(package)

    def scan_barcode(self):
        barcode = self.barcode_input.text().strip().upper()
        if barcode:
            self.scanned_barcodes.append(barcode)
            item = self.remaining_list.findItems(barcode, Qt.MatchExactly)
            if item:
                self.remaining_list.takeItem(self.remaining_list.row(item[0]))
                self.scanned_list.addItem(barcode)

            if barcode in self.df_packages['Package'].values:
                matched_row = self.df_packages[self.df_packages['Package'] == barcode].iloc[0]
                self.matched.append((barcode, matched_row['Item'], matched_row['Quantity']))
                self.result_display.append(f"Matched: {barcode} - {matched_row['Item']} - Qty: {matched_row['Quantity']}")
            else:
                self.unmatched.append(barcode)
                self.result_display.append(f"Inactive Tag: {barcode}")
            self.barcode_input.clear()
        else:
            QMessageBox.warning(self, "Input Error", "Please enter a barcode.")

    def move_item_to_scanned(self, item):
        """Move a barcode from remaining to scanned when double-clicked."""
        barcode = item.text().strip().upper()
        self.remaining_list.takeItem(self.remaining_list.row(item))
        self.scanned_list.addItem(barcode)
        self.scanned_barcodes.append(barcode)

        if barcode in self.df_packages['Package'].values:
            matched_row = self.df_packages[self.df_packages['Package'] == barcode].iloc[0]
            self.matched.append((barcode, matched_row['Item'], matched_row['Quantity']))
            self.result_display.append(f"Matched: {barcode} - {matched_row['Item']} - Qty: {matched_row['Quantity']}")
        else:
            self.unmatched.append(barcode)
            self.result_display.append(f"Inactive Tag: {barcode}")

    def move_item_to_remaining(self, item):
        """Move a barcode from scanned back to remaining when double-clicked."""
        barcode = item.text().strip().upper()
        self.scanned_list.takeItem(self.scanned_list.row(item))
        self.remaining_list.addItem(barcode)
        self.scanned_barcodes.remove(barcode)

    def search_bar_function(self):
        query = self.search_bar.text().strip().upper()
        if query:
            items = self.remaining_list.findItems(query, Qt.MatchContains)
            self.remaining_list.clearSelection()
            for item in items:
                item.setSelected(True)

    def export_pdf(self):
        if not self.scanned_barcodes:
            QMessageBox.warning(self, "No Data", "No barcodes scanned to export.")
            return

        scanned_set = set(self.scanned_barcodes)
        all_packages_set = set(self.df_packages['Package'].values)
        missing_set = all_packages_set - scanned_set
        self.missing = [
            (row['Package'], row['Item'], row['Quantity'])
            for _, row in self.df_packages[self.df_packages['Package'].isin(missing_set)].iterrows()
        ]

        output_path, _ = QFileDialog.getSaveFileName(self, "Save PDF Report", "", "PDF Files (*.pdf)")
        if output_path:
            if not output_path.endswith('.pdf'):
                output_path += '.pdf'
            try:
                export_pdf_report(self.matched, self.unmatched, self.missing, output_path)
                QMessageBox.information(self, "Success", "PDF report exported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export PDF: {e}")

    def export_to_excel_file(self):
        if not self.scanned_barcodes:
            QMessageBox.warning(self, "No Data", "No barcodes scanned to export.")
            return

        scanned_set = set(self.scanned_barcodes)
        all_packages_set = set(self.df_packages['Package'].values)
        missing_set = all_packages_set - scanned_set
        self.missing = [
            (row['Package'], row['Item'], row['Quantity'])
            for _, row in self.df_packages[self.df_packages['Package'].isin(missing_set)].iterrows()
        ]

        output_path, _ = QFileDialog.getSaveFileName(self, "Save Excel Report", "", "Excel Files (*.xlsx)")
        if output_path:
            if not output_path.endswith('.xlsx'):
                output_path += '.xlsx'
            try:
                export_to_excel(self.matched, self.unmatched, self.missing, output_path)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export Excel file: {e}")

    def clear_barcodes(self):
        self.scanned_barcodes.clear()
        self.matched.clear()
        self.unmatched.clear()
        self.missing.clear()
        self.scanned_list.clear()
        self.remaining_list.clear()
        self.result_display.clear()
        self.update_remaining_list()
        QMessageBox.information(self, "Cleared", "Scanned barcodes have been cleared.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec_())
