import os
import win32com.client
import PyPDF2
from pathlib import Path
import argparse  # import argparse

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

def main(source, destination):  # updated main function to receive source and destination as parameters
    source = os.path.abspath(source)  # get the absolute path of the source
    destination = os.path.abspath(destination)  # get the absolute path of the destination

    for dirpath, _, filenames in os.walk(source):  # use source instead of hard-coded path
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            new_dir_path = dirpath.replace(source, destination)  # use destination instead of hard-coded path
            Path(new_dir_path).mkdir(parents=True, exist_ok=True)
            
            # all extensions that can be opened by word
            doc_extensions = ['.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.odt', '.rtf', '.txt', '.wps']
            
            if filename.endswith(tuple(doc_extensions)):
                file_name_without_extension = os.path.splitext(filename)[0]
                new_file_path = os.path.join(new_dir_path, file_name_without_extension + '.pdf')
                doc_to_pdf(file_path, new_file_path)
                print("Converted: " + file_path)
            elif filename.endswith('.pdf'):
                new_file_path = os.path.join(new_dir_path, filename)
                print_pdf(file_path, new_file_path)
                print("Reprinted: " + file_path)
            else:
                new_file_path = os.path.join(new_dir_path, filename)
                print("Skipped: " + file_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()  # create an ArgumentParser object
    parser.add_argument("source", help="Source directory for the files")
    parser.add_argument("destination", help="Destination directory for the files")
    args = parser.parse_args()  # parse the arguments

    main(args.source, args.destination)  # call the main function with the parsed arguments
