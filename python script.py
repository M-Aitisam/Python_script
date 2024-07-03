import os
import pandas as pd
from docx import Document
import webbrowser

def read_excel_data(excel_file):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')  
        return df
    except FileNotFoundError:
        print(f"Error: Excel file '{excel_file}' not found.")
        return None
    except Exception as e:
        print(f"Error reading Excel file '{excel_file}': {str(e)}")
        return None

def fill_word_template(template_file, output_dir, data):
    for index, row in data.iterrows():
        name = row['Name']
        new_doc = Document(template_file)
        
        for paragraph in new_doc.paragraphs:
            if 'Dear' in paragraph.text:
                for run in paragraph.runs:
                    if 'Dear ' in run.text:
                        run.text = run.text.split('Dear ')[0] + 'Dear ' + name + ','
                        break  # Ensure we only modify once per document

        output_file_docx = os.path.join(output_dir, f"Offer_Letter_{name}.docx")
        
        new_doc.save(output_file_docx)
        webbrowser.open(output_file_docx)  # Open the generated DOCX file

def main():
    excel_file = "/home/falcon/Downloads/excelfile.xlsx"  # Excel file path
    template_file = '/home/falcon/Downloads/doc.docx'  # Word template file path
    output_dir = "/home/falcon/Downloads/pdf"  # Download directory path
    
    # Check if template file exists
    if not os.path.exists(template_file):
        print(f"Error: Template file '{template_file}' not found.")
        return

    # Check if output directory exists, create if not
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    data = read_excel_data(excel_file)
    if data is not None:
        fill_word_template(template_file, output_dir, data)
    else:
        print("Failed to read Excel data.")

if __name__ == "__main__":
    main()
