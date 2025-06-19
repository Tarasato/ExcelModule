import subprocess
import sys
import os

# --- Install required packages if not already installed ---
def install_if_missing(package_name):
    try:
        __import__(package_name)
    except ImportError:
        print(f"Installing {package_name}...")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", package_name,
            "--target", os.path.dirname(__file__)
        ])
        print(f"{package_name} installed successfully!")

install_if_missing("openpyxl")
install_if_missing("pyperclip")

# --- Imports after ensuring dependencies ---
import pyperclip
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime, time

# --- Main function ---
def highlight_missing_fields(filepath, save_path=None):
    wb = load_workbook(filepath)
    ws = wb.active

    last_row = ws.max_row
    last_col = ws.max_column
    output_col = last_col + 1  # Add column for "Missing Fields"

    ws.cell(row=1, column=output_col).value = "Missing Fields"

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green_fill = PatternFill(start_color="11FF00", end_color="11FF00", fill_type="solid")
    clear_fill = PatternFill(fill_type=None)

    # Loop through data rows
    for i in range(2, last_row + 1):
        missing_list = ""
        is_complete = True

        for j in range(1, last_col + 1):
            cell = ws.cell(row=i, column=j)
            header = ws.cell(row=1, column=j).value
            cell.fill = clear_fill

            if cell.value is None or str(cell.value).strip() == "":
                cell.fill = red_fill
                missing_list += header + ", "
                is_complete = False

        result_cell = ws.cell(row=i, column=output_col)
        result_cell.fill = clear_fill

        if is_complete:
            result_cell.value = "ALL FIELDS ARE FILLED"
            result_cell.fill = green_fill
        else:
            result_cell.value = missing_list[:-2]  # Remove last comma and space

    # Prepare output filenames
    base, ext = os.path.splitext(os.path.basename(filepath))
    new_filename_xlsx = base + "_CheckedFields" + ext
    new_filename_txt = base + "_CheckedFields.txt"

    if save_path is None:
        save_filepath_xlsx = filepath
        save_filepath_txt = os.path.splitext(filepath)[0] + ".txt"
    else:
        save_filepath_xlsx = os.path.join(save_path, new_filename_xlsx)
        save_filepath_txt = os.path.join(save_path, new_filename_txt)

    # Save updated Excel file
    wb.save(save_filepath_xlsx)
    print(f"✅ Excel saved to: {save_filepath_xlsx}")

    # Write TXT file using | without extra spacing, and formatted dates
    with open(save_filepath_txt, "w", encoding="utf-8") as txt_file:
        for i in range(1, last_row + 1):  # Include header row
            row_data = []
            for j in range(1, output_col + 1):  # Include new "Missing Fields" column
                cell = ws.cell(row=i, column=j)
                val = cell.value

                # Handle datetime formatting
                if isinstance(val, datetime):
                    if val.time() == time(0, 0):
                        formatted = val.strftime("%Y-%m-%d")
                    else:
                        formatted = val.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    formatted = str(val).strip() if val is not None else ""

                row_data.append(formatted)

            # Write row without space around pipes
            txt_file.write("|".join(row_data) + "\n")

    print(f"📄 TXT exported to: {save_filepath_txt}")

    # Copy Excel path to clipboard
    pyperclip.copy(save_filepath_xlsx)
    print("📋 Excel file path copied to clipboard!")

# --- Run function with file paths ---
highlight_missing_fields(
    r"C:\ExampleFolder\ExampleSubFolder\ExampleFile.xlsx",
    save_path=r"C:\ExampleFolder\ExampleSubFolder"
)
