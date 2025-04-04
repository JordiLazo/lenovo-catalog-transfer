import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from tkinter import ttk
import os
from dotenv import load_dotenv
import re

# Load environment variables from .env file.
load_dotenv()
# Allowed sheets for each transformation are read from the .env file.
allowed_sheet_names_CDEH = [s.strip() for s in os.getenv("CDEH", "").split(",") if s.strip()]
allowed_sheet_names_CDFI = [s.strip() for s in os.getenv("CDFI", "").split(",") if s.strip()]

def log_message(message):
    """Append a log message to the log_text widget and print to console."""
    print(message)
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)

def get_worksheet_names(file_path: Path):
    """Reads an xlsx file and returns a list of worksheet names."""
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet_names = workbook.sheetnames
        workbook.close()
        log_message(f"Loaded source file: {file_path}")
        return sheet_names
    except Exception as e:
        log_message(f"Error reading file '{file_path}': {e}")
        messagebox.showerror("Error", f"Error reading file '{file_path}': {e}")
        return []

def open_file():
    global source_file
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        log_message("No source file selected.")
        return
    source_file = Path(file_path)
    log_message(f"Source file selected: {source_file}")
    sheets = get_worksheet_names(source_file)
    update_dropdown(sheets)

def select_destination():
    """Let the user pick an existing Excel file to edit."""
    global destination_file
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        log_message("No destination file selected.")
        return
    destination_file = Path(file_path)
    log_message(f"Destination file selected: {destination_file}")
    messagebox.showinfo("Destination Selected", f"Destination file: {destination_file}")

def update_dropdown(sheets):
    sheet_list.set(sheets)
    log_message(f"Worksheets found: {sheets}")

def is_valid_row_CDEH(cell_c, cell_d, cell_e, cell_h):
    """
    Validates that for CDEH transformation:
      - Cells in columns C, D, E, and H are not empty.
      - Combined text does not contain common header keywords.
    """
    vals = [str(cell.value or "").strip() for cell in (cell_c, cell_d, cell_e, cell_h)]
    if not all(vals):
        return False
    skip_keywords = ["part no", "family - short description", "pvpr (no iva)"]
    combined = " ".join(vals).lower()
    if any(kw in combined for kw in skip_keywords):
        return False
    return True

def is_valid_row_CDFI(cell_c, cell_d, cell_f, cell_i):
    """
    Validates that for CDFI transformation:
      - Cells in columns C, D, F, and I are not empty.
      - Combined text does not contain common header keywords.
    """
    vals = [str(cell.value or "").strip() for cell in (cell_c, cell_d, cell_f, cell_i)]
    if not all(vals):
        return False
    skip_keywords = ["part no", "family - short description", "pvpr (no iva)"]
    combined = " ".join(vals).lower()
    if any(kw in combined for kw in skip_keywords):
        return False
    return True

def copy_selected_sheets():
    if not source_file or not destination_file:
        messagebox.showerror("Error", "Please select both source and destination Excel files.")
        log_message("Error: Source or destination file not selected.")
        return

    selected_indices = listbox.curselection()
    selected_sheets = [listbox.get(i) for i in selected_indices]
    log_message(f"User selected sheets: {selected_sheets}")

    # Process sheets that are in allowed lists (either group).
    sheets_to_process = [sheet for sheet in selected_sheets if (sheet in allowed_sheet_names_CDEH or sheet in allowed_sheet_names_CDFI)]
    if not sheets_to_process:
        allowed = allowed_sheet_names_CDEH + allowed_sheet_names_CDFI
        messagebox.showerror("Error", f"Please select at least one allowed sheet: {allowed}")
        log_message(f"Error: None of the selected sheets are in allowed list: {allowed}")
        return

    try:
        log_message("Starting copy operation...")
        # Open source workbook without data_only so we can access hyperlinks.
        src_wb = openpyxl.load_workbook(source_file, data_only=False)
        # Open the destination workbook.
        try:
            dest_wb = openpyxl.load_workbook(destination_file)
            log_message(f"Opened existing destination file: {destination_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open destination file: {e}")
            log_message(f"Failed to open destination file: {e}")
            src_wb.close()
            return

        dest_ws = dest_wb.active

        # Collect existing Part Numbers from destination columns C and D
        existing_part_numbers = set()
        for row in dest_ws.iter_rows(min_row=2, max_col=4, values_only=True):
            if row[0]:
                existing_part_numbers.add(str(row[0]).strip())
            if row[1]:
                existing_part_numbers.add(str(row[1]).strip())

        # Find the first empty row (assuming headers in row 1)
        dest_row = 2
        while dest_ws.cell(dest_row, 3).value is not None:
            dest_row += 1

        copied_rows = 0
        # Process each allowed sheet.
        for sheet_name in sheets_to_process:
            log_message(f"Processing sheet: {sheet_name}")
            src_ws = src_wb[sheet_name]
            # Determine transformation based on allowed group.
            if sheet_name in allowed_sheet_names_CDEH:
                # For CDEH: Use columns C (index2), D (3), E (4), H (7)
                for row_cells in src_ws.iter_rows(min_row=2):
                    cell_c = row_cells[2]
                    cell_d = row_cells[3]
                    cell_e = row_cells[4]
                    cell_h = row_cells[7]
                    if not is_valid_row_CDEH(cell_c, cell_d, cell_e, cell_h):
                        log_message("Skipping row (CDEH) due to validation failure.")
                        continue
                    part_number = str(cell_d.value or "").strip()
                    if part_number in existing_part_numbers:
                        log_message(f"Skipping duplicate Part Number (CDEH): {part_number}")
                        continue
                    val_c = str(cell_c.value or "").strip()
                    val_d = part_number
                    val_e = str(cell_e.value or "").strip()
                    val_h = cell_h.value  # could be numeric
                    merged_val = f"{val_c} - {val_e}"
                    hyperlink = cell_d.hyperlink.target if cell_d.hyperlink else ""
                    
                    dest_ws.cell(dest_row, 3, val_d)       # Destination Column C
                    dest_ws.cell(dest_row, 4, val_d)       # Destination Column D
                    dest_ws.cell(dest_row, 5, merged_val)  # Destination Column E
                    dest_ws.cell(dest_row, 8, val_h)       # Destination Column H
                    dest_ws.cell(dest_row, 10, hyperlink)  # Destination Column J
                    log_message(f"Inserted (CDEH) Part Number {val_d} at row {dest_row}")
                    existing_part_numbers.add(val_d)
                    dest_row += 1
                    copied_rows += 1
            elif sheet_name in allowed_sheet_names_CDFI:
                # For CDFI: Use columns C (index2), D (3), F (5), I (8)
                for row_cells in src_ws.iter_rows(min_row=2):
                    cell_c = row_cells[2]
                    cell_d = row_cells[3]
                    cell_f = row_cells[5]
                    cell_i = row_cells[8]
                    if not is_valid_row_CDFI(cell_c, cell_d, cell_f, cell_i):
                        log_message("Skipping row (CDFI) due to validation failure.")
                        continue
                    part_number = str(cell_d.value or "").strip()
                    if part_number in existing_part_numbers:
                        log_message(f"Skipping duplicate Part Number (CDFI): {part_number}")
                        continue
                    val_c = str(cell_c.value or "").strip()
                    val_d = part_number
                    val_f = str(cell_f.value or "").strip()
                    val_i = cell_i.value
                    merged_val = f"{val_c} - {val_f}"
                    hyperlink = cell_d.hyperlink.target if cell_d.hyperlink else ""
                    
                    dest_ws.cell(dest_row, 3, val_d)       # Destination Column C
                    dest_ws.cell(dest_row, 4, val_d)       # Destination Column D
                    dest_ws.cell(dest_row, 5, merged_val)  # Destination Column E
                    dest_ws.cell(dest_row, 8, val_i)       # Destination Column H
                    dest_ws.cell(dest_row, 10, hyperlink)  # Destination Column J
                    log_message(f"Inserted (CDFI) Part Number {val_d} at row {dest_row}")
                    existing_part_numbers.add(val_d)
                    dest_row += 1
                    copied_rows += 1

        try:
            dest_wb.save(destination_file)
            log_message(f"Successfully copied {copied_rows} rows. Duplicates were skipped.")
            messagebox.showinfo("Success", "Data copied successfully.")
        except PermissionError as pe:
            log_message(f"Permission error: {pe}. Make sure the destination file is closed in other applications.")
            messagebox.showerror("Error", f"Permission error: {pe}.")
        finally:
            src_wb.close()
            dest_wb.close()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to copy data: {e}")
        log_message(f"Error during copy: {e}")

def main():
    global sheet_list, listbox, log_text, source_file, destination_file
    source_file = None
    destination_file = None

    root = tk.Tk()
    root.title("Excel Sheet Selector")
    root.geometry("600x500")

    frame = tk.Frame(root)
    frame.pack(pady=10)

    tk.Button(frame, text="Select Excel File", command=open_file).pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Select Destination Excel", command=select_destination).pack(side=tk.LEFT, padx=5)

    tk.Label(root, text="Select Worksheet(s):").pack()
    sheet_list = tk.Variable(value=[])
    listbox = tk.Listbox(root, listvariable=sheet_list, selectmode=tk.MULTIPLE, height=6)
    listbox.pack(pady=10, fill=tk.BOTH, expand=True)

    tk.Button(root, text="Copy Data", command=copy_selected_sheets).pack(pady=10)

    tk.Label(root, text="Execution Logs:").pack()
    global log_text
    log_text = tk.Text(root, height=10, state=tk.NORMAL)
    log_text.pack(pady=10, fill=tk.BOTH, expand=True)

    root.mainloop()

if __name__ == "__main__":
    main()
