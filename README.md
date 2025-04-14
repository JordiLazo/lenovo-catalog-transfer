# Excel Data Copier - Code Explanation

This document explains the functionality and internal workings of the Python script designed to selectively copy data between Excel workbooks based on configurable rules.

## Overview

The script provides a Graphical User Interface (GUI) built with `tkinter` that allows a user to:

1.  Select a **source** Excel file (`.xlsx`).
2.  Select a **destination** Excel file (`.xlsx`).
3.  View the list of worksheets within the source file.
4.  Select one or more worksheets from the source file to process.
5.  Trigger a data copying process based on predefined rules associated with specific worksheet names.
6.  View logs of the operations performed.

The core logic focuses on reading specific columns from selected source sheets, validating rows, checking for duplicate Part Numbers in the destination, and writing processed data into the first available rows of the active sheet in the destination file.

## Key Components

1.  **Environment Variables (`.env`)**:
    *   The script uses a `.env` file to load configuration, primarily the lists of allowed worksheet names for different processing rules.
    *   Five groups are defined by environment variables: `CDEH`, `CDFI`, `CDEG`, `CDER`, `CEFH`. Each variable should contain a comma-separated string of worksheet names belonging to that group.
    *   This allows flexible configuration without modifying the code directly.

2.  **GUI (`tkinter`)**:
    *   **`main()` function**: Initializes the GUI, sets up buttons, listbox, and log area.
    *   **Buttons**:
        *   "Select Excel File": Opens a file dialog to choose the source `.xlsx` file (`open_file` function).
        *   "Select Destination Excel": Opens a file dialog to choose the destination `.xlsx` file (`select_destination` function).
        *   "Copy Data": Initiates the main data processing logic (`copy_selected_sheets` function).
    *   **Listbox (`listbox`)**: Displays worksheet names from the selected source file. Allows multiple selections.
    *   **Log Text Area (`log_text`)**: Shows status messages, errors, and progress information during execution (`log_message` function).

3.  **File Handling (`openpyxl`, `pathlib`)**:
    *   `pathlib.Path` is used for robust file path management.
    *   `openpyxl` is used to read data from the source workbook (including hyperlinks) and read/write data to the destination workbook.
    *   `get_worksheet_names`: Safely reads and returns the list of sheet names from the source file.

4.  **Data Processing Logic (`copy_selected_sheets`)**:
    *   **Input Validation**: Checks if both source and destination files have been selected.
    *   **Sheet Selection**: Retrieves the sheets selected by the user in the listbox.
    *   **Sheet Filtering**: Filters the selected sheets, processing only those whose names appear in *any* of the allowed lists defined in the `.env` file.
    *   **Workbook Loading**: Opens the source workbook (`data_only=False` to preserve hyperlinks) and the destination workbook.
    *   **Duplicate Check**:
        *   Reads columns C and D from the *destination* sheet (starting row 2).
        *   Stores all non-empty values from these columns in a `set` (`existing_part_numbers`) for efficient duplicate checking.
    *   **Destination Row Calculation**: Finds the first empty row in the destination sheet by checking for the first `None` value in column C (index 3), starting from row 2.
    *   **Row Iteration & Processing**:
        *   Iterates through each selected *and allowed* source worksheet.
        *   Determines which rule group the sheet belongs to (CDEH, CDFI, CDEG, CDER, CEFH).
        *   Iterates through rows in the source sheet (starting from row 2).
        *   **Validation**: Calls the appropriate `is_valid_row_...` function based on the sheet group. These functions check if the required cells for that group are non-empty and do not contain specific keywords (like "part no", "family - short description", "pvpr (no iva)"). Rows failing validation are skipped.
        *   **Part Number Identification**: Extracts the Part Number. This is typically from source column D, *except* for the `CEFH` group where it comes from source column E.
        *   **Duplicate Skip**: Checks if the extracted Part Number already exists in the `existing_part_numbers` set. If yes, the row is skipped.
        *   **Data Extraction & Transformation**: Extracts values from the required source columns based on the group rules. For destination column E, it typically merges values from two source columns (e.g., `f"{val_c} - {val_e}"`).
        *   **Hyperlink Extraction**: Attempts to get the hyperlink target from source column D (or potentially another column if rules were different).
        *   **Data Writing**: Writes the processed data to the calculated `dest_row` in the destination sheet according to the specific mapping:
            *   **Dest Col C**: Part Number
            *   **Dest Col D**: Part Number
            *   **Dest Col E**: Merged Description String
            *   **Dest Col H**: Specific value (from source H, I, G, R, or H depending on the group)
            *   **Dest Col J**: Hyperlink target (if found)
        *   **Update State**: Adds the newly written Part Number to the `existing_part_numbers` set and increments `dest_row`.
    *   **Saving**: Saves the modified destination workbook.
    *   **Error Handling**: Includes `try...except` blocks for file operations (like `PermissionError` if the destination file is open elsewhere) and general exceptions.
    *   **Feedback**: Updates the log and shows message boxes (success or error) to the user.

5.  **Validation Functions (`is_valid_row_...`)**:
    *   Each function (`is_valid_row_CDEH`, `is_valid_row_CDFI`, etc.) takes the relevant `openpyxl` cell objects as arguments.
    *   It checks if the `.value` of these cells (converted to string and stripped) are all non-empty.
    *   It checks if a lowercase concatenation of the cell values contains any prohibited keywords.
    *   Returns `True` if the row is valid for processing, `False` otherwise.

## Execution Flow

1.  Run the Python script.
2.  The GUI window appears.
3.  User clicks "Select Excel File" -> Chooses source file.
4.  Sheet names from the source file populate the listbox.
5.  User clicks "Select Destination Excel" -> Chooses destination file.
6.  User selects desired worksheets from the listbox.
7.  User clicks "Copy Data".
8.  The script filters selected sheets based on `.env` configuration.
9.  It reads existing part numbers from the destination.
10. It iterates through allowed sheets and their rows, applying validation and rules.
11. Valid, non-duplicate rows are processed and written to the destination file.
12. Logs are updated in the GUI.
13. The destination file is saved.
14. A final status message (success or error) is shown.
