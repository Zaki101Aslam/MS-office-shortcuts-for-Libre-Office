import os
import sys
import time
import subprocess
import shutil

# Check for required libraries inside function to allow import after install prompt
def check_deps():
    try:
        import pyautogui
        from odf import text, teletype
        from odf.opendocument import load
        return True
    except ImportError:
        print("Error: Missing required libraries.")
        print("Please install them using: pip install pyautogui odfpy")
        sys.exit(1)

def verify_shortcuts_gui():
    check_deps()
    import pyautogui
    from odf import text, teletype
    from odf.opendocument import load

    print("Starting Comprehensive GUI Verification for LibreOffice Shortcuts...")
    print("WARNING: This script will take control of your mouse and keyboard.")
    print("Please do not touch anything until it finishes.")

    # 1. Launch LibreOffice Writer
    print("\nLaunching LibreOffice Writer...")
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable:
        print("Error: LibreOffice executable not found in PATH.")
        sys.exit(1)

    proc = subprocess.Popen([executable, "--writer", "--nologo", "--nodefault"])

    # Wait for load
    print("Waiting 10 seconds for LibreOffice to load...")
    time.sleep(10)

    # Focus
    screen_width, screen_height = pyautogui.size()
    pyautogui.click(screen_width // 2, screen_height // 2)
    time.sleep(1)

    # --- Test Case 1: Backspace ---
    # Goal: Type "StartTest", Backspace 4 times, Type "Passed".
    # Result: "StartPassed"
    print("Running Test 1: Backspace...")
    pyautogui.write("StartTest", interval=0.1)
    time.sleep(0.5)
    for _ in range(4):
        pyautogui.press('backspace')
        time.sleep(0.1)
    pyautogui.write("Passed", interval=0.1)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

    # --- Test Case 2: Enter ---
    # Goal: Type "Line1", Enter, "Line2".
    # Result: Two paragraphs: "Line1", "Line2"
    print("Running Test 2: Enter...")
    pyautogui.write("Line1", interval=0.1)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.write("Line2", interval=0.1)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

    # --- Test Case 3: Delete ---
    # Goal: Type "DeleteThis", Home, Delete 6 times.
    # Result: "This"
    print("Running Test 3: Delete...")
    pyautogui.write("DeleteThis", interval=0.1)
    time.sleep(0.5)
    pyautogui.press('home')
    time.sleep(0.5)
    for _ in range(6):
        pyautogui.press('delete')
        time.sleep(0.1)
    pyautogui.press('end') # Move to end to avoid messing up next line
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

    # --- Test Case 4: Select All (Ctrl+A) and Replace ---
    # Goal: Type "SelectAll", Ctrl+A, "Replaced".
    # Result: "Replaced" (entire doc replaced? Wait, Ctrl+A selects ALL)
    # If we do Ctrl+A -> "Replaced", the ENTIRE doc becomes "Replaced".
    # This wipes out previous tests, making verification hard.
    # Better: New Document or skip Ctrl+A for now?
    # Or verify the content SO FAR, then do Ctrl+A test.
    # Let's verify the content so far by saving to a TEMP file, then continue? No, too complex.

    # Alternative: Use Ctrl+Shift+Left to select word?
    # Test Case 4: Ctrl+Shift+Left (Select Word)
    # Type "SelectWord", Ctrl+Shift+Left, "Replaced".
    # Result: "Replaced"
    print("Running Test 4: Selection (Ctrl+Shift+Left)...")
    pyautogui.write("SelectWord", interval=0.1)
    time.sleep(0.5)
    # Use standard select word shortcut
    pyautogui.hotkey('ctrl', 'shift', 'left')
    time.sleep(0.5)
    # Type overwrite
    pyautogui.write("Replaced", interval=0.1)
    time.sleep(0.5)
    pyautogui.press('enter')

    # Save
    print("Saving test results...")
    output_file = os.path.abspath("test_result.odt")
    if os.path.exists(output_file):
        os.remove(output_file)

    pyautogui.hotkey('ctrl', 's')
    time.sleep(2)
    pyautogui.write(output_file, interval=0.1)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(3)
    # Confirm overwrite if needed (might not appear if file was deleted, but safe to press)
    pyautogui.press('enter')
    time.sleep(1)

    # Close
    print("Closing LibreOffice...")
    pyautogui.hotkey('ctrl', 'q')
    time.sleep(2)
    if proc.poll() is None:
        proc.terminate()

    # Verify
    print("\nVerifying Document Content...")
    if not os.path.exists(output_file):
        print(f"FAILED: Output file {output_file} was not created.")
        sys.exit(1)

    try:
        doc = load(output_file)
        paragraphs = []
        for p in doc.getElementsByType(text.P):
            content = teletype.extractText(p).strip()
            if content: # Ignore empty lines
                paragraphs.append(content)

        print(f"Read Paragraphs: {paragraphs}")

        # Expected paragraphs based on tests:
        # 1. "StartPassed"
        # 2. "Line1"
        # 3. "Line2"
        # 4. "This"
        # 5. "Replaced"

        expected = ["StartPassed", "Line1", "Line2", "This", "Replaced"]

        # Filter paragraphs to remove empty strings if any
        paragraphs = [p for p in paragraphs if p]

        errors = []
        if len(paragraphs) != len(expected):
            errors.append(f"Paragraph count mismatch: Expected {len(expected)}, Got {len(paragraphs)}")
            errors.append(f"Expected: {expected}")
            errors.append(f"Got: {paragraphs}")
        else:
            for i, (got, exp) in enumerate(zip(paragraphs, expected)):
                if got != exp:
                    errors.append(f"Line {i+1}: Expected '{exp}', Got '{got}'")

        if errors:
            print("FAILURE: Verification Failed!")
            for e in errors:
                print(f"  - {e}")
            sys.exit(1)
        else:
            print("SUCCESS: All key binding tests passed!")

    except Exception as e:
        print(f"Error analyzing document: {e}")
        sys.exit(1)

if __name__ == "__main__":
    verify_shortcuts_gui()
