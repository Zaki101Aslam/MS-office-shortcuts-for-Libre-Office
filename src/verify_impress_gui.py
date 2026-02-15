from test_utils import launch_app, save_and_close
import pyautogui
import time
import os
import sys
from odf import draw, teletype
from odf.opendocument import load

def verify_impress_gui():
    print("Starting Impress GUI Verification...")
    proc = launch_app("impress")

    # 1. New Slide (Ctrl+M)
    print("Test: New Slide (Ctrl+M)")
    # Usually starts with 1 slide.
    pyautogui.hotkey('ctrl', 'm')
    time.sleep(1)
    # Should be on Slide 2.

    # 2. Duplicate Slide (Ctrl+Shift+D)
    print("Test: Duplicate Slide (Ctrl+Shift+D)")
    pyautogui.hotkey('ctrl', 'shift', 'd')
    time.sleep(1)
    # Should be on Slide 3 (copy of Slide 2).

    # 3. Add Text
    print("Test: Type Text on Slide 3")
    # Click to focus slide content? Usually typing works if slide is selected?
    # Impress selection model is tricky.
    # Often need to click "Click to add title".
    # Let's try F2 (Edit Text)
    pyautogui.press('f2')
    time.sleep(0.5)
    pyautogui.write("Slide3Text", interval=0.1)
    pyautogui.press('esc')

    output_file = os.path.abspath("test_impress.odp")
    save_and_close(output_file, proc)

    # Verify Content
    print("Verifying content...")
    doc = load(output_file)
    slides = doc.getElementsByType(draw.Page)

    print(f"Slide Count: {len(slides)}")

    # Expect: 3 slides (1 Initial + 1 New + 1 Duplicate)
    if len(slides) < 3:
        print(f"FAILURE: Expected at least 3 slides, got {len(slides)}. (Ctrl+M or Ctrl+Shift+D failed)")
        sys.exit(1)

    # Check text on the last slide
    # Text in Impress is in draw:frame -> draw:text-box -> text:p
    last_slide = slides[-1]
    slide_text = ""
    for frame in last_slide.getElementsByType(draw.Frame):
        slide_text += teletype.extractText(frame).strip()

    print(f"Text on last slide: '{slide_text}'")

    if "Slide3Text" not in slide_text:
        # F2 might not have worked if focus was on slide pane.
        # But if we verify slide count, we verified the shortcut logic at least partially.
        # But let's fail if text missing.
        print("FAILURE: Text 'Slide3Text' not found on last slide.")
        sys.exit(1)

    print("SUCCESS: Impress Verification Passed.")

if __name__ == "__main__":
    verify_impress_gui()
