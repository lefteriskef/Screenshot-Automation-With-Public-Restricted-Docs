import keyboard
import pyautogui
import pyperclip
from docx import Document
from docx.shared import Inches
import time
import os
from datetime import datetime
import sys
import threading

docxfile = r"C:\Users\lefte\Downloads\Screenshots_project\Snap-Save-URL-Screenshot-Automation(2-file-option)\shots_public.docx"
docxfile2 = r"C:\Users\lefte\Downloads\Screenshots_project\Snap-Save-URL-Screenshot-Automation(2-file-option)\shots_restricted.docx"
shotfile = r"C:\Users\lefte\Downloads\Screenshots_project\Snap-Save-URL-Screenshot-Automation(2-file-option)\shots\shot.png"

# Ensure files exist and load as two separate Document objects
if not os.path.exists(docxfile):
    doc_public = Document()
    doc_public.save(docxfile)
else:
    doc_public = Document(docxfile)

if not os.path.exists(docxfile2):
    doc_restricted = Document()
    doc_restricted.save(docxfile2)
else:
    doc_restricted = Document(docxfile2)

exit_flag = threading.Event()  # Flag to signal exit


def do_cap1():
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    shot = pyautogui.screenshot()
    shot.save(shotfile)
    doc_public.add_paragraph(f"Screenshot taken at {timestamp}:")
    doc_public.add_picture(shotfile, width=Inches(7))   # screenshot taken

    keyboard.press_and_release('ctrl+l')
    time.sleep(1)
    keyboard.press_and_release('ctrl+c')
    time.sleep(1)
    url = pyperclip.paste()
    print(f"--------------Captured URL: {url}")
    doc_public.add_paragraph(f"Captured URL: {url}")
    doc_public.save(docxfile)
    print('Screenshot and URL saved and appended to PUBLIC docx.')


def do_cap2():
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    shot = pyautogui.screenshot()
    shot.save(shotfile)
    doc_restricted.add_paragraph(f"Screenshot taken at {timestamp}:")
    doc_restricted.add_picture(shotfile, width=Inches(7))   # screenshot taken

    keyboard.press_and_release('ctrl+l')
    time.sleep(1)
    keyboard.press_and_release('ctrl+c')
    time.sleep(1)
    url = pyperclip.paste()
    print(f"--------------Captured URL: {url}")
    doc_restricted.add_paragraph(f"Captured URL: {url}")
    doc_restricted.save(docxfile2)
    print('Screenshot and URL saved and appended to RESTRICTED docx.')


def input_listener():
    while not exit_flag.is_set():
        user_input = input()
        if user_input.strip().lower() == "exit":
            print("Exit command received. Exiting script.")
            exit_flag.set()
            keyboard.unhook_all_hotkeys()


keyboard.add_hotkey('ctrl+q', do_cap1)
keyboard.add_hotkey('ctrl+y', do_cap2)

print("Script running. Press Ctrl+Q to take screenshot and grab URL (public).")
print("Script running. Press Ctrl+Y to take screenshot and grab URL (restricted).")
print("Type 'exit' (without quotes) and press Enter to exit safely.")

threading.Thread(target=input_listener, daemon=True).start()

# Replace keyboard.wait() with loop, checking exit_flag
while not exit_flag.is_set():
    time.sleep(0.1)  # Short sleep to reduce CPU use

print("Cleanup done, script exiting.")
sys.exit(0)