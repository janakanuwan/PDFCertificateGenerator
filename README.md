# PDFCertificateGenerator

This script automates the generation of PDF certificates based on data from an Excel file and a PowerPoint template.
The certificates are saved as PDF files.

## Requirements

- Microsoft PowerPoint is installed
- **Python Version**: Python 3.8 or later
- **Python Libraries**:
    - `pandas`
    - `python_pptx`
    - `datetime` (built-in with Python)
    - `comtypes`
    - `openpyxl`
- **Input Files**:
    - An Excel file named `Recipients.xlsx`
    - A PowerPoint template named `Certificate_Template.pptx`

## Installation Guide

### Step 1: Clone or Download the Script

Download the script and place it in your working directory along with the `Recipients.xlsx`
and `Certificate_Template.pptx` files.

### Step 2: Create a Virtual Environment (Optional but Recommended)

Run the following commands to create and activate a virtual environment:

#### On macOS/Linux:

```bash
python -m venv certificate
source certificate/bin/activate
```

### Step 3: Install Required Libraries

Install the dependencies using `pip`:

```
pip install pandas python_pptx openpyxl comtypes
```

### Step 4: Prepare Input Files

Ensure the following files are in the same directory as the script:

- Excel File: `Recipients.xlsx`
    - Contains the data for generating recommendation letters (Sheet name: `Details`).
- Word Template: `Certificate_Template.pptx`
    - Contains placeholders like `[[Name]]`, `[[Date]]`, etc.

### Step 5: Run the Script

Execute the script using the following command:

```
python generate_certificates.py
```

Resulting Word and PDF files will be inside "GeneratedCertificates" folder.

## References

- Some codes are generated with the help of ChatGPT