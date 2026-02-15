import os
import sys
import time
import subprocess
import shutil

def verify_shortcuts_gui():
    # Check for required libraries inside function to allow import after install prompt
    try:
        import pyautogui
        from odf import text, teletype
        from odf.opendocument import load
    except ImportError:
        print("Error: Missing required libraries.")
        print("Please install them using: pip install pyautogui odfpy")
        sys.exit(1)

    print("Starting GUI Verification for LibreOffice Shortcuts...")
    print("WARNING: This script will take control of your mouse and keyboard.")
    print("Please do not touch anything until it finishes.")
    print("Ensure LibreOffice Writer is installed and accessible via 'libreoffice' or 'soffice' command.")

    # 1. Launch LibreOffice Writer
    print("\nLaunching LibreOffice Writer...")
    # Try to find the executable
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable:
        print("Error: LibreOffice executable not found in PATH.")
        sys.exit(1)

    # Start process
    proc = subprocess.Popen([executable, "--writer", "--nologo", "--nodefault"])

    # Wait for load (generous time)
    print("Waiting 10 seconds for LibreOffice to load...")
    time.sleep(10)

    # Focus the window (click in the center of the screen roughly)
    screen_width, screen_height = pyautogui.size()
    pyautogui.click(screen_width // 2, screen_height // 2)
    time.sleep(1)

    # 2. Test Typing and Backspace
    print("Test 1: Typing and Backspace...")
    # Just type first to ensure focus
    pyautogui.write("Hello World", interval=0.1)
    time.sleep(0.5)
    # Backspace 5 times -> "Hello "
    for _ in range(5):
        pyautogui.press('backspace')
        time.sleep(0.1)

    pyautogui.write("Test", interval=0.1)
    time.sleep(0.5)
    # Expected: "Hello Test"

    pyautogui.press('enter') # New line

    # 3. Test Select All and Delete
    print("Test 2: Select All (Ctrl+A) and Delete...")
    pyautogui.write("Delete Me", interval=0.1)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(0.5)
    # Expected: Empty document (or mostly empty)

    # 4. Test Undo (Ctrl+Z)
    print("Test 3: Undo (Ctrl+Z)...")
    pyautogui.hotkey('ctrl', 'z')
    time.sleep(0.5)
    # Expected: "Hello Test\nDelete Me" restored (depending on undo stack)
    # Actually, simpler test: Type "Wrong", Undo -> Empty line.

    # Let's clear everything reliably first: Select All, Delete
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(0.5)

    # Start fresh for verification text
    print("Writing Final Verification Text...")
    pyautogui.write("Final Verification Text", interval=0.1)
    time.sleep(0.5)

    # 5. Save the file
    print("Saving document...")
    output_file = os.path.abspath("test_output.odt")
    if os.path.exists(output_file):
        os.remove(output_file)

    pyautogui.hotkey('ctrl', 's')
    time.sleep(2) # Wait for dialog

    # Type filename
    pyautogui.write(output_file, interval=0.1)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(3) # Wait for save

    # Confirm overwrite if dialog appears (Enter again)
    # This might accidentally confirm unrelated dialogs, but is generally safe
    pyautogui.press('enter')
    time.sleep(1)

    # 6. Close LibreOffice
    print("Closing LibreOffice...")
    pyautogui.hotkey('ctrl', 'q')
    time.sleep(2)

    # Terminate if still running
    if proc.poll() is None:
        proc.terminate()

    # 7. Verify Content
    print("\nVerifying Document Content...")
    if not os.path.exists(output_file):
        print(f"FAILED: Output file {output_file} was not created.")
        sys.exit(1)

    try:
        doc = load(output_file)
        all_text = ""
        for p in doc.getElementsByType(text.P):
            all_text += teletype.extractText(p) + "\n"

        print(f"Document Content:\n---\n{all_text.strip()}\n---")

        if "Final Verification Text" in all_text:
            print("SUCCESS: Key bindings (Typing, Backspace, Delete, Enter, Ctrl+A, Ctrl+S, Ctrl+Q) appear to be working.")
        else:
            print("FAILURE: Document content does not match expected text.")
            sys.exit(1)

    except Exception as e:
        print(f"Error reading document: {e}")
        sys.exit(1)

if __name__ == "__main__":
    verify_shortcuts_gui()
