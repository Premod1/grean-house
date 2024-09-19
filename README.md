# Excel to PDF Converter with PyQt5

This project is a desktop application built using PyQt5 that allows users to upload an Excel file, generate an HTML table with formatted data, and convert it into a downloadable PDF. The application is intuitive and user-friendly, streamlining the process of generating professional-looking PDFs from Excel data.

## Features

- **Upload Excel File**: Supports Excel files with specific columns (`Item`, `Invoice`, `Name`, `Price`, and `City`).
- **Generate HTML Table**: Automatically creates an HTML table based on the data in the Excel file, with custom formatting.
- **Convert HTML to PDF**: Uses `pdfkit` to convert the generated HTML into a downloadable PDF.
- **User Notifications**: Provides success or error alerts during the file upload, HTML creation, and PDF download processes.
- **Simple, Clean UI**: Built with PyQt5, the application has an easy-to-use graphical interface.

## Requirements

- Python 3.x
- PyQt5
- pandas
- openpyxl
- pdfkit
- wkhtmltopdf (for PDF conversion)

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/Premod1/grean-house.git

   ```

2. Install the required Python packages:

   ```bash
    pip install PyQt5 pandas openpyxl pdfkit

   ```

3. Install wkhtmltopdf for your operating system:
   ```bash
   sudo apt-get install wkhtmltopdf

## Usage

   Run the Python application:

```bash
    python main.py

Use the GUI to upload an Excel file containing the required columns (Item, Invoice, Name, Price, City).

Once the file is uploaded, the app generates an HTML document with a formatted table and provides an option to download it as a PDF.

The generated PDF will be saved in your Downloads folder with a timestamped filename.

File Structure
plaintext
Copy code
.
├── main.py # Main application script
├── README.md # This README file
├── requirements.txt # List of required packages
└── final_output.html # Generated HTML file (created during runtime)
