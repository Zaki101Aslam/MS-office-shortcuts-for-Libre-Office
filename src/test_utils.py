import os
import sys
import time
import subprocess
import shutil

# Check for required libraries inside function to allow import after install prompt
def check_deps():
    try:
        import pyautogui
        from odf import text, teletype, table, draw
        from odf.opendocument import load
        return True
    except ImportError:
        print("Error: Missing required libraries.")
        print("Please install them using: pip install pyautogui odfpy")
        sys.exit(1)

def get_executable():
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable:
        print("Error: LibreOffice executable not found in PATH.")
        sys.exit(1)
    return executable

def launch_app(app_type):
    """
    app_type: 'writer', 'calc', or 'impress'
    """
    check_deps()
    import pyautogui

    print(f"\nLaunching LibreOffice {app_type.capitalize()}...")
    executable = get_executable()

    args = [executable, f"--{app_type}", "--nologo", "--nodefault"]
    proc = subprocess.Popen(args)

    # Wait for load
    print("Waiting 10 seconds for LibreOffice to load...")
    time.sleep(10)

    # Focus
    screen_width, screen_height = pyautogui.size()
    pyautogui.click(screen_width // 2, screen_height // 2)
    time.sleep(1)

    return proc

def save_and_close(output_file, proc):
    import pyautogui

    print(f"Saving to {output_file}...")
    if os.path.exists(output_file):
        os.remove(output_file)

    pyautogui.hotkey('ctrl', 's')
    time.sleep(2)
    pyautogui.write(output_file, interval=0.1)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(3)
    # Confirm overwrite if needed
    pyautogui.press('enter')
    time.sleep(1)

    print("Closing LibreOffice...")
    pyautogui.hotkey('ctrl', 'q')
    time.sleep(2)

    if proc.poll() is None:
        proc.terminate()

    if not os.path.exists(output_file):
        print(f"FAILED: Output file {output_file} was not created.")
        sys.exit(1)

    print(f"File {output_file} created successfully.")
