
# PDF to Excel Data Extraction Script

This Python script extracts financial data from a specific page of a PDF document and appends the extracted data to an Excel file. Follow the steps below to use the script.

## Requirements

Ensure you have the following Python libraries installed:

- `PyPDF2` for reading PDF files
- `openpyxl` for working with Excel files

Install these libraries using pip if you haven't already:

```bash
pip install PyPDF2 openpyxl
```

## Script Overview

The script performs the following tasks:

1. **Reads a specific page from a PDF file.**
2. **Extracts relevant financial data from the text of that page.**
3. **Appends the extracted data to an Excel file, creating the file if it does not exist.**

## Python Script

Save the following code to a file named `extract_data.py`:

```python
import PyPDF2
import re
import openpyxl

def read_specific_page(file_path, page_number):
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            if page_number < 1 or page_number > len(pdf_reader.pages):
                return f"Page {page_number} does not exist in the PDF."
            page = pdf_reader.pages[page_number - 1]
            text = page.extract_text()
            return text
    except FileNotFoundError:
        return f"File {file_path} not found."
    except Exception as e:
        return str(e)

def extract_savings_info(text):
    savings_potential_pattern = r"Savings potential\s+(\d{1,2}\.\d+%)"
    mt_co2_savings_pattern = r"([\d,]+)\s*MT CO 2 per year savings"
    mt_fuel_savings_pattern = rf"(\d+)\s*\n\s*{re.escape('Main Engine')}"
    roi_pattern = r"Return of investment\s+\(years\)\s+(\d+\.\d+)"
  
    savings_potential = re.search(savings_potential_pattern, text)
    mt_co2_savings = re.search(mt_co2_savings_pattern, text)
    mt_fuel_savings = re.search(mt_fuel_savings_pattern, text)
    roi = re.search(roi_pattern, text)
  
    return {
        "savings_potential": savings_potential.group(1) if savings_potential else None,
        "mt_co2_savings": mt_co2_savings.group(1).replace(',', '') if mt_co2_savings else None,
        "mt_fuel_savings": mt_fuel_savings.group(1).replace(',', '') if mt_fuel_savings else None,
        "roi": roi.group(1) if roi else None
    }

def append_data_excel(data):
    try:
        # Load workbook, create if it doesn't exist
        try:
            book = openpyxl.load_workbook('data.xlsx')
        except FileNotFoundError:
            book = openpyxl.Workbook()
      
        active = book.active
      
        # Append headers if the sheet is empty
        if active.max_row == 1:
            active.append(['Savings Potential', 'MT CO2 Savings', 'MT Fuel Savings', 'ROI'])
      
        # Append the data
        active.append([
            data.get('savings_potential'),
            data.get('mt_co2_savings'),
            data.get('mt_fuel_savings'),
            data.get('roi')
        ])
      
        # Save the workbook
        book.save('data.xlsx')
    except Exception as e:
        print(f"An error occurred while appending data to Excel: {e}")

if __name__ == '__main__':
    file_path = 'sample.pdf'  # Replace with your PDF file path
    page_number = 8  # Specify the page number you want to read
    text = read_specific_page(file_path, page_number)
  
    if "Page" in text or "File" in text:  # Error messages if any
        print(text)
    else:
        data = extract_savings_info(text)
        append_data_excel(data)
```

## Usage Instructions

1. **Replace File Path**:

   - Update the `file_path` variable in the script with the path to your PDF file.
2. **Specify Page Number**:

   - Set the `page_number` variable to the page number you want to extract data from.
3. **Run the Script**:

   - Execute the script using Python:

     ```bash
     python extract_data.py
     ```
4. **Check Excel File**:

   - The script will create or update `data.xlsx` with the extracted data. Open this file to view the appended data.

## Troubleshooting

- **File Not Found**: Ensure the PDF file path is correct.
- **No Data Extracted**: Verify that the patterns in the `extract_savings_info` function match the content in your PDF.

---

This Markdown file provides a clear guide for users to run the Python script, including setup instructions, code, and troubleshooting tips.
