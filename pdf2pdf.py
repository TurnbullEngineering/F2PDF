from pywinauto.keyboard import send_keys
import pywinauto
import os
import time
import argparse


def reprint_pdf(
    source_path, destination_path, app: pywinauto.application.Application = None
):
    """
    Prints the PDF file to the destination path

    source_path: Path to the source PDF file
    destination_path: Path to the destination PDF file, directory must exist
    app: pywinauto.application.Application object, will be created if not provided, should be provided to reduce overhead
    """
    # TODO: Gracefully close the application and all of its child windows if an error occurs
    if not app:
        app = pywinauto.application.Application().start(
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
        )

    try:
        desktop = pywinauto.Desktop(backend="uia")
        time.sleep(0.5)

        # Open the file
        send_keys("^o")
        time.sleep(0.5)  # wait for the dialog to open

        # Dialog where we enter the file name
        open_dialog = app.window(title="Open")
        open_dialog.Edit.set_text(source_path)
        time.sleep(0.5)
        # Press enter to open the file
        send_keys("{ENTER}")

        # Document name without extension
        document_name = os.path.splitext(os.path.basename(source_path))[0]
        # Get the main window
        main_window = desktop.window(
            title_re=f"{os.path.basename(document_name)}.*- Adobe Acrobat.*",
            class_name="AcrobatSDIWindow",
        )
        main_window.set_focus()

        # Use keyboard shortcut to print to destination
        send_keys("^p")
        time.sleep(2)

        print_dialog = main_window.child_window(title="Print", control_type="Window")

        print_dialog.Print.click()
        time.sleep(1)

        ms_print_dialog = desktop.window(
            title="Save Print Output As", top_level_only=False
        )

        # Enter the destination path
        ms_print_dialog.child_window(title="File name:", control_type="Edit").set_text(
            destination_path
        )

        # Save the file
        ms_print_dialog.Save.click()
        time.sleep(0.5)

        progress_dialog = desktop.window(
            title="Progress", top_level_only=False, control_type="Window"
        )

        while progress_dialog.exists():
            time.sleep(0.5)

        print(f"\033[92mProcessed: {source_path}\033[0m")
    except Exception as e:
        print(f"\033[91mAn error occurred while processing {source_path}: {e}\033[0m")
    finally:
        app.kill()
    time.sleep(0.5)
    return True


def main(source, destination):
    source = os.path.abspath(source)
    destination = os.path.abspath(destination)

    reprint_pdf(source, destination)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Copy protected PDF files")
    parser.add_argument("source", help="Source path for the PDF file")
    parser.add_argument("destination", help="Destination path for the PDF file")
    args = parser.parse_args()

    main(args.source, args.destination)
