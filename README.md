# Lenovo catalog transfer

A Python tool that helps extract and consolidate data from multiple Excel worksheets into a single destination file according to specific formatting rules. The program uses a graphical interface for easy selection of source files, destination files, and worksheets.

## Overview

The program provides a Graphical User Interface (GUI) built with `tkinter` that allows a user to:

1.  Select a **source** Excel file (`.xlsx`).
2.  Select a **destination** Excel file (`.xlsx`).
3.  View the list of worksheets within the source file.
4.  Select one or more worksheets from the source file to process.
5.  Trigger a data copying process based on predefined rules associated with specific worksheet names in a specific Excel defined in a format.
6.  View logs of the operations performed.

The core logic focuses on reading specific columns from selected source sheets, validating rows, checking for duplicate Part Numbers in the destination, and writing processed data into the first available rows of the active sheet in the destination file.

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
