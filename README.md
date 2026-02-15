# MS Office Shortcuts for LibreOffice

This repository provides LibreOffice configuration files that map common Microsoft Office keyboard shortcuts to their LibreOffice equivalents. It covers Writer (Word), Calc (Excel), and Impress (PowerPoint).

## Features

*   **Word -> Writer:** Maps shortcuts like Clear Formatting (`Ctrl+Space`), Go To (`Ctrl+G`), Font Size (`Ctrl+Shift+<`, `>`), etc.
*   **Excel -> Calc:** Maps shortcuts like Insert/Delete Rows (`Ctrl++`, `Ctrl+-`), Insert Date/Time (`Ctrl+;`, `Ctrl+Shift+:`), etc.
*   **PowerPoint -> Impress:** Maps shortcuts like New Slide (`Ctrl+M`), Duplicate Slide (`Ctrl+D`), Presentation Mode (`F5`, `Shift+F5`).
*   **Customization Tool:** Includes a Python script to easily modify shortcuts and generate your own configuration files.

## Installation

### Using Pre-generated Configurations

1.  Download the `.cfg` file for the application you want to configure from the `dist/` folder:
    *   `Word_Shortcuts_for_Writer.cfg`
    *   `Excel_Shortcuts_for_Calc.cfg`
    *   `PowerPoint_Shortcuts_for_Impress.cfg`

2.  Open LibreOffice (Writer, Calc, or Impress).

3.  Go to **Tools** > **Customize...**

4.  Select the **Keyboard** tab.

5.  Click the **Load...** button on the right side.

6.  Select the downloaded `.cfg` file.

7.  Click **Open**.

8.  Verify the shortcuts are loaded and click **OK**.

## Customization

If you want to change any of the mappings or add new ones, you can use the included Python script.

### Prerequisites

*   Python 3 installed.

### Steps

1.  Clone this repository.
    ```bash
    git clone https://github.com/Zaki101Aslam/MS-office-shortcuts-for-Libre-Office
    cd MS-office-shortcuts-for-Libre-Office
    ```

2.  Run the customization script in interactive mode:
    ```bash
    python3 src/generate_config.py --interactive
    ```

3.  Follow the on-screen prompts:
    *   Select the application (Writer, Calc, or Impress).
    *   **Edit (E):** Change an existing shortcut.
    *   **Add (A):** Add a new shortcut mapping (requires knowing the UNO command).
    *   **Save (S):** Save your changes and generate a new `.cfg` file.

4.  Import the generated `.cfg` file into LibreOffice as described in the Installation section.

## Contributing

Feel free to open issues or pull requests to suggest more mappings! The mappings are stored in JSON files in the `mappings/` directory.
