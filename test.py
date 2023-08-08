from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import time


def print_identifiers():
    app = Application().start(r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe")
    time.sleep(2)

    send_keys("^o")
    time.sleep(2)

    print("--- Main Window ---")
    main_window = app.window()
    print(main_window.print_control_identifiers(2))


print_identifiers()
