from docx import Document
from docx2pdf import convert
import os

def replace_text_in_word(doc_path, replacements):
    # Load the document
    doc = Document(doc_path)

    # Replace text according to the dictionary provided
    for paragraph in doc.paragraphs:
        for old_text, new_text in replacements.items():
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text)

    # Save the modified document
    new_doc_path = doc_path[-16:]
    if new_doc_path.startswith(' '):
        new_doc_path = new_doc_path.lstrip()
    doc.save(new_doc_path)

    return new_doc_path

def convert_docx_to_pdf(doc_path):
    # Convert the Word document to PDF
    pdf_path = doc_path.replace('.docx', '.pdf')
    convert(doc_path, pdf_path)
    return pdf_path

def process_document(doc_path, name, reg_no):
    # Define the replacements
    replacements = {
        "Name-": f"Name- {name}",
        "Reg.no-": f"Reg.no- {reg_no}"
    }

    # Replace text in the document
    modified_doc_path = replace_text_in_word(doc_path, replacements)

    # Convert the modified document to PDF
    pdf_path = convert_docx_to_pdf(modified_doc_path)

    return pdf_path

def process_all_documents_in_directory(directory):
    name_input = input("Enter the Name: ")
    reg_no_input = input("Enter the Reg.no: ")
    # Loop through all files in the directory    
    for filename in os.listdir(directory):
        if filename.endswith('.docx'):
            doc_path = os.path.join(directory, filename)
            print(f"Processing: {doc_path}")
            pdf_output = process_document(doc_path, name_input, reg_no_input)
            print(f"PDF created: {pdf_output}")

# Get user input for Name and Reg.no


# Example usage with a dynamic directory path
directory_path = os.getcwd()  # Replace with your directory path
process_all_documents_in_directory(directory_path)
