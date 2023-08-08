import os
import win32com.client
import pywinauto
from pathlib import Path
import argparse
import time

from pdf2pdf import reprint_pdf


def copy_word_document(source_path, destination_path, word):
    """Opens original document, opens new document, copies original document to new document, saves new document"""
    try:
        source_doc = word.Documents.Open(source_path)

        # Wait for the document to open
        while source_doc.ReadOnly:
            time.sleep(0.1)

        time.sleep(1)

        source_doc.Content.Select()
        word.Selection.Copy()

        # Create a new document
        destination_doc = word.Documents.Add()

        # Wait for the new document to be ready
        while destination_doc.ReadOnly:
            time.sleep(0.5)

        time.sleep(1.5)

        destination_doc.Content.Paste()

        # Wait for the content to be pasted
        time.sleep(1)

        # Save the new document to the destination path
        destination_doc.SaveAs(FileName=destination_path, FileFormat=16)

        # Wait for the document to be saved
        while not destination_doc.Saved:
            time.sleep(1)
            destination_doc.SaveAs(FileName=destination_path, FileFormat=16)

        # Close the documents
        source_doc.Close(SaveChanges=0)
        destination_doc.Close()
        # success in green
        # print("Processed: " + source_path)
        print(f"\033[92mProcessed: {source_path}\033[0m")

    except Exception as e:
        # f"An error occurred while processing {source_path}: {e}"
        # error in red
        print(f"\033[91mAn error occurred while processing {source_path}: {e}\033[0m")
        # Dump info about the error to the console in red
        if source_doc:
            source_doc.Close(SaveChanges=0)
        if destination_doc:
            destination_doc.Close(SaveChanges=0)
        exit()


def main(
    source, destination
):  # updated main function to receive source and destination as parameters
    source = os.path.abspath(source)  # get the absolute path of the source
    destination = os.path.abspath(
        destination
    )  # get the absolute path of the destination

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    time.sleep(1)

    for dirpath, _, filenames in os.walk(source):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            new_dir_path = dirpath.replace(source, destination)
            Path(new_dir_path).mkdir(parents=True, exist_ok=True)

            # all extensions that can be opened by word
            doc_extensions = [".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm"]

            if filename.endswith(tuple(doc_extensions)):
                file_name_without_extension = os.path.splitext(filename)[0]
                new_file_path = os.path.join(
                    new_dir_path, file_name_without_extension + ".docx"
                )
                try:
                    copy_word_document(file_path, new_file_path, word)
                except Exception as e:
                    print(f"Error while copying {file_path}: {e}")

            elif filename.endswith(".pdf"):
                new_file_path = os.path.join(new_dir_path, filename)
                try:
                    reprint_pdf(file_path, new_file_path)
                    time.sleep(1)
                except Exception as e:
                    print(f"\033[91mError while reprinting {file_path}: {e}\033[0m")

            else:
                new_file_path = os.path.join(new_dir_path, filename)
                # print("Skipped: " + file_path)
                time.sleep(0.1)

    # Close Word without keeping changes (just in case)
    word.Quit(SaveChanges=0)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()  # create an ArgumentParser object
    parser.add_argument("source", help="Source directory for the files")
    parser.add_argument("destination", help="Destination directory for the files")
    args = parser.parse_args()  # parse the arguments

    main(
        args.source, args.destination
    )  # call the main function with the parsed arguments
