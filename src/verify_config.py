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
