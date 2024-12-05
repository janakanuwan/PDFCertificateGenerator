# RecommendationLetterGenerator

This script automates the generation of personalized recommendation letters based on data from an Excel file and a Word
template. The letters are saved as PDF files.

## Requirements

- Microsoft Word is installed
- **Python Version**: Python 3.8 or later
- **Python Libraries**:
    - `pandas`
    - `python-docx`
    - `datetime` (built-in with Python)
    - `docx2pdf`
    - `openpyxl`
- **Input Files**:
    - An Excel file named `Applications-details.xlsx`
    - A Word template named `RecommendationLetter_Template.docx`

## Installation Guide

### Step 1: Clone or Download the Script

Download the script and place it in your working directory along with the `Applications-details.xlsx`
and `RecommendationLetter_Template.docx` files.

### Step 2: Create a Virtual Environment (Optional but Recommended)

Run the following commands to create and activate a virtual environment:

#### On macOS/Linux:

```bash
python -m venv recommendation
source recommendation/bin/activate
```

### Step 3: Install Required Libraries

Install the dependencies using `pip`:

```
pip install pandas python-docx docx2pdf openpyxl
```

### Step 4: Prepare Input Files

Ensure the following files are in the same directory as the script:

- Excel File: `Applications-details.xlsx`
    - Contains the data for generating recommendation letters (Sheet name: `Details`).
- Word Template: `RecommendationLetter_Template.docx`
    - Contains placeholders like `<Date>`, `<Recommendation-Committee>`, etc.

### Step 5: Run the Script

Execute the script using the following command:

```
python generate_letters.py
```

Resulting Word and PDF files will be inside "GeneratedLetters" folder.

## References

- Some codes are generated with the help of ChatGPT