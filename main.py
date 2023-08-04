import os
import win32com.client
import PyPDF2
from pathlib import Path

def doc_to_pdf(doc_path, pdf_path):
    """Convert Word Document to PDF"""
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

def print_pdf(pdf_path, new_file_path):
    """Print PDF document"""
    reader = PyPDF2.PdfReader(pdf_path)
    writer = PyPDF2.PdfWriter()

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        writer.add_page(page)

    with open(new_file_path, 'wb') as new_file:
        writer.write(new_file)

def main():
    root_dir = r'C:\Users\VinhDinh\Repos\PF2PDF\from'  # Update this to your directory
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            new_dir_path = dirpath.replace(root_dir, r'C:\Users\VinhDinh\Repos\PF2PDF\to')  # Update this to your new directory
            Path(new_dir_path).mkdir(parents=True, exist_ok=True)
            
            if filename.endswith('.doc') or filename.endswith('.docx'):
                new_file_path = os.path.join(new_dir_path, filename.replace('.docx', '.pdf').replace('.doc', '.pdf'))
                doc_to_pdf(file_path, new_file_path)
            elif filename.endswith('.pdf'):
                new_file_path = os.path.join(new_dir_path, filename)
                print_pdf(file_path, new_file_path)

if __name__ == "__main__":
    main()