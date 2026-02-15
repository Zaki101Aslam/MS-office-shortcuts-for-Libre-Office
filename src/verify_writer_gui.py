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

    # Logic is slightly complex because Headers might appear out of order if we iterate types separately.
    # Let's iterate all children of body.text
    all_content = []
    for element in doc.text.childNodes:
        if element.qname == (text.P.qname):
            t = teletype.extractText(element).strip()
            if t: all_content.append(t)
        elif element.qname == (text.H.qname):
            t = teletype.extractText(element).strip()
            if t: all_content.append(f"H: {t}")

    print(f"Ordered Content: {all_content}")

    expected = [
        "BackPass",
        "H: HeadingText",
        "Correct",
        "Gone",
        "BoldText"
    ]

    errors = []
    if len(all_content) != len(expected):
        errors.append(f"Length mismatch: Expected {len(expected)}, Got {len(all_content)}")

    for i, (got, exp) in enumerate(zip(all_content, expected)):
        if got != exp:
             errors.append(f"Line {i+1}: Expected '{exp}', Got '{got}'")

    if errors:
        print("FAILURE: Writer Verification Failed.")
        for e in errors: print(f"  - {e}")
        sys.exit(1)
    else:
        print("SUCCESS: Writer Verification Passed.")

if __name__ == "__main__":
    verify_writer_gui()
