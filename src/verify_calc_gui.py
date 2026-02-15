from test_utils import launch_app, save_and_close
import pyautogui
import time
import os
import sys
from odf import table, teletype
from odf.opendocument import load

def verify_calc_gui():
    print("Starting Calc GUI Verification...")
    proc = launch_app("calc")

    # 1. Navigation & Data Entry
    print("Test: Data Entry & Arrows")
    # A1
    pyautogui.write("ValA1", interval=0.1)
    pyautogui.press('right')
    time.sleep(0.5)
    # B1
    pyautogui.write("ValB1", interval=0.1)
    pyautogui.press('enter')
    # Usually moves down to B2? Or A2? Depends on settings.
    # Default Calc: Down. So now at B2.
    time.sleep(0.5)

    # 2. Home Key
    print("Test: Home Key")
    pyautogui.press('home') # Should go to A2
    time.sleep(0.5)
    pyautogui.write("ValA2", interval=0.1)
    pyautogui.press('right') # B2
    time.sleep(0.5)

    # 3. Fill Down (Ctrl+D)
    print("Test: Fill Down (Ctrl+D)")
    # At B2. Type "FillSource".
    pyautogui.write("FillSource", interval=0.1)
    pyautogui.press('enter') # B3
    time.sleep(0.5)
    pyautogui.press('up') # Back to B2
    time.sleep(0.5)
    # Select B2:B3. Shift+Down
    pyautogui.hotkey('shift', 'down')
    time.sleep(0.5)
    # Ctrl+D
    pyautogui.hotkey('ctrl', 'd')
    time.sleep(0.5)

    # 4. New Sheet (Shift+F11)
    print("Test: New Sheet (Shift+F11)")
    pyautogui.hotkey('shift', 'f11')
    time.sleep(1)

    # Handle potential "Insert Sheet" dialog (Calc sometimes asks for name/position)
    # Pressing Enter confirms default (which is usually OK)
    pyautogui.press('enter')
    time.sleep(1)

    # Should be on Sheet 2 (or new sheet)
    pyautogui.write("Sheet2Data", interval=0.1)
    pyautogui.press('enter')

    output_file = os.path.abspath("test_calc.ods")
    save_and_close(output_file, proc)

    # Verify Content
    print("Verifying content...")
    doc = load(output_file)
    sheets = doc.getElementsByType(table.Table)

    if len(sheets) < 2:
        print(f"FAILURE: Expected at least 2 sheets, got {len(sheets)}")
        sys.exit(1)

    # Verify Sheet 1 (usually the second one in list if inserted before? Or after? Shift+F11 inserts *before* usually in Excel, but Calc?)
    # Let's check all sheets for our data.

    content_map = {}
    for sheet in sheets:
        sheet_name = sheet.getAttribute("name")
        rows = sheet.getElementsByType(table.TableRow)
        sheet_data = []
        for row in rows:
            row_data = []
            cells = row.getElementsByType(table.TableCell)
            for cell in cells:
                # Cells can be repeated.
                repeat = int(cell.getAttribute("numbercolumnsrepeated") or 1)
                text_content = teletype.extractText(cell).strip()
                for _ in range(repeat):
                    row_data.append(text_content)
            sheet_data.append(row_data)
        content_map[sheet_name] = sheet_data

    # We expect "Sheet2Data" in one sheet
    sheet2_found = False
    for s_name, data in content_map.items():
        if data and len(data) > 0 and data[0] and data[0][0] == "Sheet2Data":
            sheet2_found = True
            print(f"Found Sheet2 data in {s_name}")

    if not sheet2_found:
        print("FAILURE: 'Sheet2Data' not found in any sheet (New Sheet shortcut failed?)")
        sys.exit(1)

    # We expect A1="ValA1", B1="ValB1", A2="ValA2", B2="FillSource", B3="FillSource"
    # Note: Sheet ordering might vary. Look for the sheet containing "ValA1".
    sheet1_found = False
    for s_name, data in content_map.items():
        # Check A1
        if len(data) > 0 and len(data[0]) > 0 and data[0][0] == "ValA1":
            sheet1_found = True
            print(f"Found Sheet1 data in {s_name}")
            # Check B1
            if len(data[0]) < 2 or data[0][1] != "ValB1":
                print(f"FAILURE: B1 mismatch. Got {data[0][1] if len(data[0])>1 else 'None'}")
                sys.exit(1)
            # Check A2
            if len(data) < 2 or len(data[1]) < 1 or data[1][0] != "ValA2":
                print(f"FAILURE: A2 mismatch. Got {data[1][0] if len(data)>1 else 'None'}")
                sys.exit(1)
            # Check B3 (Filled)
            if len(data) < 3 or len(data[2]) < 2 or data[2][1] != "FillSource":
                print(f"FAILURE: B3 (Fill Down) mismatch. Got {data[2][1] if len(data)>2 and len(data[2])>1 else 'None'}")
                sys.exit(1)
            break

    if not sheet1_found:
        print("FAILURE: Sheet1 data not found.")
        sys.exit(1)

    print("SUCCESS: Calc Verification Passed.")

if __name__ == "__main__":
    verify_calc_gui()
