from test_utils import launch_app, save_and_close
import pyautogui
import time
import os
import sys
from odf import text, teletype
from odf.opendocument import load

def verify_writer_gui():
    print("Starting Writer GUI Verification...")
    proc = launch_app("writer")

    # 1. Backspace
    print("Test: Backspace")
    pyautogui.write("BackTest", interval=0.1)
    time.sleep(0.5)
    for _ in range(4):
        pyautogui.press('backspace')
        time.sleep(0.1)
    pyautogui.write("Pass", interval=0.1)
    pyautogui.press('enter')
    # Expect: "BackPass"

    # 2. Styles (Ctrl+1 Heading 1)
    print("Test: Styles (Ctrl+1)")
    pyautogui.write("HeadingText", interval=0.1)
    pyautogui.hotkey('ctrl', '1') # Heading 1
    pyautogui.press('enter')

    # 3. Undo/Redo
    print("Test: Undo/Redo (Ctrl+Z/Y)")
    pyautogui.write("Mistake", interval=0.1)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'z') # Undo "Mistake"
    time.sleep(0.5)
    pyautogui.write("Correct", interval=0.1) # Write "Correct"
    pyautogui.press('enter')
    # Expect: "Correct" (after previous line)

    # 4. Selection & Delete
    print("Test: Selection (Ctrl+Shift+Left) & Delete")
    pyautogui.write("DelWord", interval=0.1)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'shift', 'left')
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(0.5)
    pyautogui.write("Gone", interval=0.1)
    pyautogui.press('enter')
    # Expect: "Gone"

    # 5. Bold (Ctrl+B)
    print("Test: Bold (Ctrl+B)")
    pyautogui.hotkey('ctrl', 'b')
    pyautogui.write("BoldText", interval=0.1)
    pyautogui.hotkey('ctrl', 'b') # Toggle off
    pyautogui.press('enter')

    output_file = os.path.abspath("test_writer.odt")
    save_and_close(output_file, proc)

    # Verify Content
    print("Verifying content...")
    doc = load(output_file)
    paragraphs = []
    for p in doc.getElementsByType(text.P):
        content = teletype.extractText(p).strip()
        if content:
            paragraphs.append(content)

    # Headers are distinct from paragraphs in ODF text extraction?
    # teletype.extractText usually extracts all.
    # Heading 1 might be H element?
    for h in doc.getElementsByType(text.H):
        content = teletype.extractText(h).strip()
        if content:
             paragraphs.append(f"HEADER: {content}")

    print(f"Read Content: {paragraphs}")

    # Simplify verification: Check presence of expected text in the document
    # Order might be preserved by getElementsByType if we just look at the list we already built
    # The 'paragraphs' list above already contains Ps and Hs? No, I added Hs to it manually?

    # Re-collect everything robustly
    all_text_content = []

    # Collect Paragraphs
    for p in doc.getElementsByType(text.P):
        t = teletype.extractText(p).strip()
        if t: all_text_content.append(t)

    # Collect Headers
    for h in doc.getElementsByType(text.H):
        t = teletype.extractText(h).strip()
        if t: all_text_content.append(f"H: {t}")

    print(f"Collected Content: {all_text_content}")

    # We expect these strings to be present. Exact order is hard without robust traversal,
    # but let's verify existence which proves the shortcuts worked.

    expected_items = {
        "BackPass": "Backspace Test",
        "H: HeadingText": "Heading Style Test",
        "Correct": "Undo/Redo Test",
        "Gone": "Selection/Delete Test",
        "BoldText": "Bold Test"
    }

    errors = []
    for exp, desc in expected_items.items():
        if exp not in all_text_content:
            errors.append(f"Missing expected content: '{exp}' ({desc})")

if __name__ == "__main__":
    verify_writer_gui()
