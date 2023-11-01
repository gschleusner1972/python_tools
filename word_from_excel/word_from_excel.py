import os
from openpyxl import load_workbook
from docx import Document

# Determine the current directory and navigate to the root directory "python_tools\word_from_excel"
CURRENT_DIR = os.path.abspath(os.getcwd())
BASE_DIR = os.path.join(CURRENT_DIR.split("python_tools")[0], "python_tools", "word_from_excel")

def replace_placeholder_with_value(doc, placeholder, value):
    """Replace all occurrences of a placeholder in a Word document with a given value."""
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                run.text = run.text.replace(placeholder, value)
                
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.text = run.text.replace(placeholder, value)
                            
    return doc

def create_documents_from_template(excel_file, template_file, output_folder):
    # Load the Excel workbook and select the active sheet
    wb = load_workbook(excel_file)
    ws = wb.active
    
    # Get headers from the first row to identify placeholders
    headers = [cell.value for cell in ws[1]]
    
    # Ensure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Iterate over each record in the Excel file (skipping the header)
    for row in ws.iter_rows(min_row=2, values_only=True):
        doc = Document(template_file)
        
        # For each header, replace its corresponding placeholder with the value from the current row
        for header, value in zip(headers, row):
            replace_placeholder_with_value(doc, f"[[{header}]]", str(value))
        
        # Save the modified document in the output folder
        output_filename = os.path.join(output_folder, f"output_{row[0]}.docx")
        doc.save(output_filename)

    print(f"Generated {ws.max_row - 1} documents in the '{output_folder}' directory.")

if __name__ == "__main__":
    # Using the base directory for all paths
    create_documents_from_template(
        os.path.join(BASE_DIR, "data.xlsx"),
        os.path.join(BASE_DIR, "template.docx"),
        os.path.join(BASE_DIR, "output")
    )
