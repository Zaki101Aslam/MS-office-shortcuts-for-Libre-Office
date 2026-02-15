import json
import os
import zipfile
import argparse
import sys
import xml.sax.saxutils

# Key Mappings
KEY_MAP = {
    "SPACE": "KEY_SPACE",
    "ENTER": "KEY_RETURN",
    "RETURN": "KEY_RETURN",
    "ESC": "KEY_ESCAPE",
    "ESCAPE": "KEY_ESCAPE",
    "BACKSPACE": "KEY_BACKSPACE",
    "DELETE": "KEY_DELETE",
    "UP": "KEY_UP",
    "DOWN": "KEY_DOWN",
    "LEFT": "KEY_LEFT",
    "RIGHT": "KEY_RIGHT",
    "HOME": "KEY_HOME",
    "END": "KEY_END",
    "PAGEUP": "KEY_PAGEUP",
    "PAGEDOWN": "KEY_PAGEDOWN",
    "TAB": "KEY_TAB",
    "INSERT": "KEY_INSERT",
    "+": "KEY_ADD",      # Numpad +
    "PLUS": "KEY_ADD",   # Alias
    "-": "KEY_SUBTRACT", # Numpad -
    "MINUS": "KEY_SUBTRACT",
    "=": "KEY_EQUAL",
    "EQUAL": "KEY_EQUAL",
    ".": "KEY_POINT",
    "POINT": "KEY_POINT",
    ",": "KEY_COMMA",
    "COMMA": "KEY_COMMA",
    ";": "KEY_SEMICOLON",
    "SEMICOLON": "KEY_SEMICOLON",
    ":": "KEY_SEMICOLON", # Shift handled separately
    "[": "KEY_BRACKETLEFT", # Check specific code
    "]": "KEY_BRACKETRIGHT",
    "BRACKETLEFT": "KEY_BRACKETLEFT",
    "BRACKETRIGHT": "KEY_BRACKETRIGHT",
    "'": "KEY_QUOTELEFT", # Verify
    "QUOTE": "KEY_QUOTELEFT",
    "<": "KEY_LESS",
    ">": "KEY_GREATER",
}

# Populate standard keys
for char in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789":
    KEY_MAP[char] = f"KEY_{char}"
for i in range(1, 13):
    KEY_MAP[f"F{i}"] = f"KEY_F{i}"

def parse_shortcut(shortcut_str):
    # Handle Ctrl++ case where split("+") creates empty strings
    # "Ctrl++" -> ["Ctrl", "", ""]
    # "Ctrl+Shift++" -> ["Ctrl", "Shift", "", ""]
    parts = shortcut_str.upper().split("+")

    modifiers = {
        "shift": "false",
        "mod1": "false", # Ctrl
        "mod2": "false"  # Alt
    }

    code = None

    # Filter empty parts which represent the literal "+" key (if it was a separator)
    # But wait, "A+B" -> ["A", "B"]. "A++" -> ["A", "", ""].
    # If we have empty strings, it means there were adjacent + signs.
    # "Ctrl++" means Ctrl and +.

    cleaned_parts = []
    i = 0
    while i < len(parts):
        part = parts[i]
        if part == "":
            # Found a gap, meaning we hit a + separator that was followed by another +
            # or preceded by one?
            # "Ctrl++" -> split -> "Ctrl", "", ""
            # The key is likely "+"
            cleaned_parts.append("+")
            # Skip the next empty string if it exists?
            # "Ctrl++".split("+") is ['Ctrl', '', '']
            # The last empty string is because + is at the end.
            i += 1
        else:
            cleaned_parts.append(part)
        i += 1

    # Re-process cleaned parts
    final_parts = []
    for part in cleaned_parts:
        if part == "CTRL":
            modifiers["mod1"] = "true"
        elif part == "ALT":
            modifiers["mod2"] = "true"
        elif part == "SHIFT":
            modifiers["shift"] = "true"
        elif part == "":
            pass # ignore
        else:
            final_parts.append(part)

    if len(final_parts) == 0:
        # Maybe it was just "+" ?
        if "+" in shortcut_str and "CTRL" not in shortcut_str.upper():
             code = KEY_MAP["+"]
    elif len(final_parts) == 1:
        key = final_parts[0]
        if key in KEY_MAP:
            code = KEY_MAP[key]
        else:
            code = f"KEY_{key}"
    else:
        # Multiple non-modifier keys? Error.
        # But wait, maybe the loop logic above missed something.
        # Let's trust that the last part is the key.
        key = final_parts[-1]
        if key in KEY_MAP:
            code = KEY_MAP[key]
        else:
            code = f"KEY_{key}"

    return code, modifiers

def create_xml(mappings):
    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<!DOCTYPE accel:acceleratorlist PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "accelerator.dtd">',
        '<accel:acceleratorlist xmlns:accel="http://openoffice.org/2001/accel" xmlns:xlink="http://www.w3.org/1999/xlink">'
    ]

    for mapping in mappings:
        shortcut = mapping["ms_shortcut"]
        command = mapping["uno_command"]

        code, modifiers = parse_shortcut(shortcut)

        if not code:
            print(f"Warning: Could not parse key for {shortcut}")
            continue

        # Escape special characters in XML attributes
        command_escaped = xml.sax.saxutils.escape(command)
        code_escaped = xml.sax.saxutils.escape(code)

        attr_str = f'accel:code="{code_escaped}" xlink:href="{command_escaped}"'
        if modifiers["shift"] == "true":
            attr_str += ' accel:shift="true"'
        if modifiers["mod1"] == "true":
            attr_str += ' accel:mod1="true"'
        if modifiers["mod2"] == "true":
            attr_str += ' accel:mod2="true"'

        lines.append(f' <accel:item {attr_str}/>')

    lines.append('</accel:acceleratorlist>')
    return "\n".join(lines)

def create_manifest():
    return """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/>
 <manifest:file-entry manifest:full-path="Configurations2/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/>
 <manifest:file-entry manifest:full-path="Configurations2/accelerator/current.xml" manifest:media-type=""/>
</manifest:manifest>"""

def generate_package(json_path, output_path, defaults_path=None):
    print(f"Reading {json_path}...")
    with open(json_path, 'r') as f:
        custom_mappings = json.load(f)

    # Load defaults if provided
    final_mappings = []
    if defaults_path and os.path.exists(defaults_path):
        print(f"Reading defaults from {defaults_path}...")
        with open(defaults_path, 'r') as f:
            default_mappings = json.load(f)

        # Create a dictionary of custom shortcuts to easily check for overrides
        # Keying by shortcut string (normalized) to detect conflicts
        custom_keys = {m["ms_shortcut"].upper().replace(" ", ""): m for m in custom_mappings}

        # Add defaults ONLY if they don't conflict with custom mappings
        for default in default_mappings:
            key = default["ms_shortcut"].upper().replace(" ", "")
            if key not in custom_keys:
                final_mappings.append(default)
            else:
                print(f"Overriding default {default['ms_shortcut']} with custom {custom_keys[key]['ms_shortcut']}")

        # Add all custom mappings
        final_mappings.extend(custom_mappings)
    else:
        final_mappings = custom_mappings

    xml_content = create_xml(final_mappings)
    manifest_content = create_manifest()

    with zipfile.ZipFile(output_path, 'w') as zf:
        # Mimetype should be first and uncompressed
        zf.writestr("mimetype", "application/vnd.sun.xml.ui.configuration", compress_type=zipfile.ZIP_STORED)
        zf.writestr("Configurations2/accelerator/current.xml", xml_content)
        zf.writestr("META-INF/manifest.xml", manifest_content)

    print(f"Generated {output_path}")

def interactive_mode():
    print("Interactive Mode")
    print("----------------")

    files = {
        "1": ("mappings/writer.json", "Writer"),
        "2": ("mappings/calc.json", "Calc"),
        "3": ("mappings/impress.json", "Impress")
    }

    print("Select application:")
    for k, v in files.items():
        print(f"{k}. {v[1]}")

    choice = input("Choice: ")
    if choice not in files:
        print("Invalid choice.")
        return

    filepath, app_name = files[choice]

    if not os.path.exists(filepath):
        print(f"Mapping file {filepath} not found.")
        return

    with open(filepath, 'r') as f:
        mappings = json.load(f)

    while True:
        print(f"\nCurrent Mappings for {app_name}:")
        for idx, m in enumerate(mappings):
            print(f"{idx + 1}. {m['command_name']} -> {m['ms_shortcut']} ({m['uno_command']})")

        print("\nOptions: (E)dit, (A)dd, (S)ave & Exit, (Q)uit")
        opt = input("Option: ").upper()

        if opt == "Q":
            break
        elif opt == "S":
            with open(filepath, 'w') as f:
                json.dump(mappings, f, indent=4)
            print("Saved mappings.")

            gen = input("Generate .cfg file now? (Y/n): ").upper()
            if gen != "N":
                # Determine output path based on app_name
                if not os.path.exists("dist"):
                    os.makedirs("dist")

                out_path = f"dist/Word_Shortcuts_for_Writer.cfg"
                def_path = "defaults/writer.json"

                if app_name == "Calc":
                    out_path = f"dist/Excel_Shortcuts_for_Calc.cfg"
                    def_path = "defaults/calc.json"
                elif app_name == "Impress":
                    out_path = f"dist/PowerPoint_Shortcuts_for_Impress.cfg"
                    def_path = "defaults/impress.json"

                generate_package(filepath, out_path, def_path)
            break
        elif opt == "E":
            idx = int(input("Enter number to edit: ")) - 1
            if 0 <= idx < len(mappings):
                item = mappings[idx]
                print(f"Editing {item['command_name']}")
                new_shortcut = input(f"New Shortcut (current: {item['ms_shortcut']}): ")
                if new_shortcut:
                    item['ms_shortcut'] = new_shortcut
            else:
                print("Invalid index.")
        elif opt == "A":
            name = input("Command Name: ")
            uno = input("UNO Command (e.g. .uno:Save): ")
            sc = input("Shortcut (e.g. Ctrl+S): ")
            mappings.append({
                "command_name": name,
                "uno_command": uno,
                "ms_shortcut": sc
            })

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate LibreOffice shortcut config")
    parser.add_argument("--map", help="Path to JSON mapping file")
    parser.add_argument("--out", help="Output path for .cfg file")
    parser.add_argument("--interactive", action="store_true", help="Run in interactive mode")

    args = parser.parse_args()

    if args.interactive:
        interactive_mode()
    elif args.map and args.out:
        generate_package(args.map, args.out)
    else:
        # Default behavior: generate all
        if not os.path.exists("dist"):
            os.makedirs("dist")

        if os.path.exists("mappings/writer.json"):
            generate_package("mappings/writer.json", "dist/Word_Shortcuts_for_Writer.cfg", "defaults/writer.json")
        if os.path.exists("mappings/calc.json"):
            generate_package("mappings/calc.json", "dist/Excel_Shortcuts_for_Calc.cfg", "defaults/calc.json")
        if os.path.exists("mappings/impress.json"):
            generate_package("mappings/impress.json", "dist/PowerPoint_Shortcuts_for_Impress.cfg", "defaults/impress.json")
