import pandas as pd
from docx import Document
from datetime import datetime
from docx2pdf import convert
import os

# Constants
EXCEL_FILE = "Applications-details.xlsx"
EXCEL_SHEET = "Details"
WORD_TEMPLATE = "RecommendationLetter_Template.docx"
OUTPUT_DIR = "GeneratedLetters"

# Step 1: Load the Excel File
try:
    data = pd.read_excel(EXCEL_FILE, sheet_name=EXCEL_SHEET)
    print("Excel file loaded successfully.")
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Step 2: Filter rows with non-empty ID
data = data.dropna(subset=["ID"])
if data.empty:
    print("No valid rows found in the Excel sheet.")
    exit()

# Step 3: Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Step 4: Process each row and generate recommendation letters
for _, row in data.iterrows():
    try:
        # Extract data
        recommendation_id = row.get("ID", None)
        committee = row.get("Recommendation-Committee", None)
        position = row.get("Recommendation-Position", "")
        department = row.get("Recommendation-Department", "")

        print(
            f"\nProcessing row ID: [{recommendation_id}], Committee: [{committee}], Position: [{position}], Department: [{department}]"
        )

        # Check for NaN values and handle them explicitly
        if (pd.isna(recommendation_id) or pd.isna(committee) or
                not str(recommendation_id).strip() or not str(committee).strip()):
            print(f"Skipping row due to empty ID or Committee")
            continue

        # Convert values to strings and strip whitespace or handle NaN
        recommendation_id = "" if pd.isna(recommendation_id) else str(recommendation_id).strip()
        committee = "" if pd.isna(committee) else str(committee).strip()
        position = "" if pd.isna(position) else str(position).strip()
        department = "" if pd.isna(department) else str(department).strip()

        # Open the Word template
        try:
            doc = Document(WORD_TEMPLATE)
        except Exception as e:
            print(f"Error opening Word template: {e}")
            continue

        # Replace placeholders in the document
        for paragraph in doc.paragraphs:
            if "<Date>" in paragraph.text:  # Format: Mon, DD, YYY
                paragraph.text = paragraph.text.replace("<Date>", datetime.now().strftime("%b %d, %Y"))
            if "<Recommendation-Committee>" in paragraph.text:
                paragraph.text = paragraph.text.replace("<Recommendation-Committee>", committee)
            if "<Recommendation-Position>" in paragraph.text:
                paragraph.text = paragraph.text.replace("<Recommendation-Position>", position)
            if "<Recommendation-Department>" in paragraph.text:
                paragraph.text = paragraph.text.replace("<Recommendation-Department>", department)

        # Save the personalized Word document
        word_file = os.path.join(OUTPUT_DIR, f"{recommendation_id}.docx")
        doc.save(word_file)

        # Convert to PDF
        try:
            convert(word_file, os.path.join(OUTPUT_DIR, f"{recommendation_id}.pdf"))
        except Exception as e:
            print(f"Error converting {word_file} to PDF: {e}")
            continue

        # Remove the intermediate Word file
        # os.remove(word_file)

        print(f"Generated PDF for ID: {recommendation_id}")

    except KeyError as e:
        print(f"Missing expected column in row: {e}")
        exit()
    except Exception as e:
        print(f"Error processing row ID {row['ID']}: {e}")

print("\nProcessing complete.")
