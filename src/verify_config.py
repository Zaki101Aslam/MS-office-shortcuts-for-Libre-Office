import os
import zipfile
import xml.etree.ElementTree as ET
import glob
import sys

def verify_cfg(cfg_path):
    print(f"Verifying {cfg_path}...")
    errors = []

    try:
        with zipfile.ZipFile(cfg_path, 'r') as zf:
            if "Configurations2/accelerator/current.xml" not in zf.namelist():
                errors.append("Missing Configurations2/accelerator/current.xml")
                return errors

            xml_content = zf.read("Configurations2/accelerator/current.xml")

            try:
                root = ET.fromstring(xml_content)

                # Check namespace
                # ElementTree handles namespaces like {http://openoffice.org/2001/accel}item
                ns = {'accel': 'http://openoffice.org/2001/accel', 'xlink': 'http://www.w3.org/1999/xlink'}

                # Use ElementTree iteration to handle namespaces more robustly
                items = root.findall('.//{http://openoffice.org/2001/accel}item')

                seen_keys = set()

                for item in items:
                    code = item.get('{http://openoffice.org/2001/accel}code')
                    shift = item.get('{http://openoffice.org/2001/accel}shift')
                    mod1 = item.get('{http://openoffice.org/2001/accel}mod1')
                    mod2 = item.get('{http://openoffice.org/2001/accel}mod2')
                    command = item.get('{http://www.w3.org/1999/xlink}href')

                    if not code:
                        errors.append(f"Item missing code: {ET.tostring(item)}")
                        continue

                    if not command:
                        errors.append(f"Item missing command: {ET.tostring(item)}")
                        continue

                    # Validate Command Structure
                    if not command.startswith(".uno:"):
                        errors.append(f"Invalid command format (must start with .uno:): {command}")

                    # Check for duplicates
                    key_sig = f"{code}_{shift}_{mod1}_{mod2}"
                    if key_sig in seen_keys:
                        errors.append(f"Duplicate key assignment: {code} (shift={shift}, mod1={mod1}, mod2={mod2}) assigned to {command}")
                    seen_keys.add(key_sig)

                # Safety Check: Verify presence of critical keys
                CRITICAL_KEYS = {
                    "KEY_BACKSPACE": "Backspace",
                    "KEY_DELETE": "Delete",
                    "KEY_RETURN": "Enter",
                    "KEY_ESCAPE": "Escape",
                    "KEY_TAB": "Tab",
                    "KEY_UP": "Up Arrow",
                    "KEY_DOWN": "Down Arrow",
                    "KEY_LEFT": "Left Arrow",
                    "KEY_RIGHT": "Right Arrow",
                    "KEY_HOME": "Home",
                    "KEY_END": "End",
                    "KEY_PAGEUP": "Page Up",
                    "KEY_PAGEDOWN": "Page Down"
                }

                # Helper function to check if a key is present (regardless of modifiers for base keys, but checking specific combos for others)
                def find_key(code, shift=None, mod1=None):
                    for k in seen_keys:
                        parts = k.split('_')
                        # Format: CODE_SHIFT_MOD1_MOD2
                        # We just need to match the CODE part for basic keys like Arrows/Enter/Backspace
                        # (Usually these don't require modifiers to simply function, though shift+arrow selects)
                        # Actually, we want to ensure at least the BASE version exists (no modifiers).
                        # key_sig = f"{code}_{shift}_{mod1}_{mod2}"

                        k_code = "_".join(parts[:-3]) # Handle codes with underscores? No, codes don't have underscores usually.
                        # Wait, KEY_PAGE_UP has underscores.
                        # The splitting logic above is flawed if codes have underscores.
                        # Better approach: store tuples in seen_keys or parse differently.
                        # But seen_keys stores strings. Let's re-parse.
                        pass

                    # Better: Iterate through items again or build a structured set
                    pass

                # Let's rebuild a structured view of present keys
                present_codes = set()
                present_shortcuts = set() # Stores (code, shift, mod1, mod2)

                for item in items:
                    c = item.get('{http://openoffice.org/2001/accel}code')
                    s = item.get('{http://openoffice.org/2001/accel}shift') == 'true'
                    m1 = item.get('{http://openoffice.org/2001/accel}mod1') == 'true'
                    m2 = item.get('{http://openoffice.org/2001/accel}mod2') == 'true'
                    if c:
                        present_codes.add(c)
                        present_shortcuts.add((c, s, m1, m2))

                for key_code, key_name in CRITICAL_KEYS.items():
                    # check for base key (no modifiers)
                    if (key_code, False, False, False) not in present_shortcuts:
                        errors.append(f"CRITICAL MISSING: {key_name} ({key_code}) is not bound!")

                # Check for critical Ctrl shortcuts (Copy, Cut, Paste, Undo, Redo, Save, Select All)
                CTRL_SHORTCUTS = {
                    "KEY_C": "Ctrl+C (Copy)",
                    "KEY_X": "Ctrl+X (Cut)",
                    "KEY_V": "Ctrl+V (Paste)",
                    "KEY_Z": "Ctrl+Z (Undo)",
                    "KEY_Y": "Ctrl+Y (Redo)",
                    "KEY_S": "Ctrl+S (Save)",
                    "KEY_A": "Ctrl+A (Select All)"
                }

                for key_code, name in CTRL_SHORTCUTS.items():
                    # Check for Ctrl+Key (mod1=True, others=False)
                    if (key_code, False, True, False) not in present_shortcuts:
                         errors.append(f"CRITICAL MISSING: {name} is not bound!")

            except ET.ParseError as e:
                errors.append(f"XML Parse Error: {e}")

    except zipfile.BadZipFile:
        errors.append("Invalid Zip File")
    except Exception as e:
        errors.append(f"Unexpected error: {e}")

    return errors

if __name__ == "__main__":
    if not os.path.exists("dist"):
        print("dist/ directory not found.")
        sys.exit(1)

    cfg_files = glob.glob("dist/*.cfg")
    if not cfg_files:
        print("No .cfg files found in dist/")
        sys.exit(1)

    all_passed = True
    for cfg in cfg_files:
        errors = verify_cfg(cfg)
        if errors:
            all_passed = False
            print(f"FAILED: {cfg}")
            for err in errors:
                print(f"  - {err}")
        else:
            print(f"PASSED: {cfg}")

    if not all_passed:
        sys.exit(1)
    else:
        print("\nAll configurations passed verification.")
