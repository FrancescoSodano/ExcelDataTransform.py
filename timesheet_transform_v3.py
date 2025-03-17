import os
import sys
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import timedelta

def build_records_from_timesheet(timesheet_file, mapping_file):
    """
    1) Reads the timesheet Excel, which may have multiple sheets.
       - If a sheet lacks a 'Codice Commessa' column, uses the sheet name.
       - Reads 'WeekRange', 'Autore', etc. to build day-by-day records.
    2) Reads the mapping file to map original 'Codice Commessa' → 'COMMESSA'.
    3) Returns a DataFrame aggregated by [DATA, SURNAME] with joined COMMESSA and summed ORE.
    """

    # ------------------------
    # Read the mapping
    # ------------------------
    df_map = pd.read_excel(mapping_file)
    commessa_map = dict(zip(df_map.iloc[:, 0], df_map.iloc[:, 1]))

    # ------------------------
    # Read all sheets of timesheet
    # ------------------------
    sheets_dict = pd.read_excel(timesheet_file, sheet_name=None)
    df_raw_list = []
    for sheet_name, df_sheet in sheets_dict.items():
        if "Codice Commessa" not in df_sheet.columns:
            df_sheet["Codice Commessa"] = sheet_name
        df_raw_list.append(df_sheet)
    df_raw = pd.concat(df_raw_list, ignore_index=True)

    # ------------------------
    # Weekday offset (Italian)
    # ------------------------
    day_offset = {
        "Lunedì": 0,
        "Martedì": 1,
        "Mercoledì": 2,
        "Giovedì": 3,
        "Venerdì": 4,
        "Sabato": 5,
        "Domenica": 6
    }

    records = []
    for idx, row in df_raw.iterrows():
        week_range_str = str(row.get("WeekRange", "")).strip()
        if " al " not in week_range_str:
            continue
        try:
            start_str, _ = week_range_str.split(" al ")
            start_date = pd.to_datetime(start_str, dayfirst=True)
        except Exception as e:
            print(f"Skipping row {idx} due to date parse error: {e}")
            continue

        # Map commessa
        codice_commessa = row.get("Codice Commessa", "")
        commessa_final = commessa_map.get(codice_commessa, codice_commessa)

        # Autore → surname
        autore = str(row.get("Autore", "")).strip()
        surname = autore.split()[-1] if autore else "UNKNOWN"

        # Check each weekday column for hours
        for day_col, offset in day_offset.items():
            hours = row.get(day_col, 0)
            if pd.notna(hours) and float(hours) != 0.0:
                actual_date = start_date + timedelta(days=offset)
                records.append({
                    "DATA": actual_date.date(),   # or keep as a Timestamp
                    "COMMESSA": commessa_final,
                    "ORE": float(hours),
                    "SURNAME": surname
                })

    if not records:
        print("No valid timesheet records found.")
        return pd.DataFrame()  # empty

    df_timesheet = pd.DataFrame(records)

    # ------------------------
    # Aggregate by [DATA, SURNAME]
    # Combine commesse (unique) and sum ORE
    # ------------------------
    df_agg = df_timesheet.groupby(["DATA", "SURNAME"], as_index=False).agg({
        "COMMESSA": lambda x: "; ".join(sorted(set(x))),
        "ORE": "sum"
    })
    # Convert DATA to datetime for matching
    df_agg["DATA"] = pd.to_datetime(df_agg["DATA"])
    return df_agg


def update_strategie_in_place(strategie_file, df_agg):
    """
    For each sheet in 'strategie_file' (sheet = surname):
      - We assume columns A, B, C = [DATA, COMMESSA, ORE].
      - For each row in that sheet, if the date is found in df_agg for that surname,
        update the COMMESSA and ORE cell. Preserve existing cell formatting.
      - If the date isn't found, we do nothing (we don't create new rows).
    """
    # Load the workbook with openpyxl (preserves formatting, tables, styles)
    wb = openpyxl.load_workbook(strategie_file)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Filter df_agg for this surname
        subset = df_agg[df_agg["SURNAME"] == sheet_name].copy()
        if subset.empty:
            continue
        # Index by date for quick lookups
        subset.set_index("DATA", inplace=True)

        # Iterate the rows of the sheet; assume row 1 is headers: A=DATA, B=COMMESSA, C=ORE
        for row_cells in ws.iter_rows(min_row=2, max_col=3, values_only=False):
            date_cell, commessa_cell, ore_cell = row_cells  # openpyxl cell objects

            # Convert the cell's value to a date
            cell_value = date_cell.value
            if isinstance(cell_value, str):
                try:
                    cell_value = pd.to_datetime(cell_value).date()
                except:
                    continue
            elif hasattr(cell_value, "date"):
                cell_value = cell_value.date()

            # If this date is in subset index, update COMMESSA and ORE
            if cell_value in subset.index.date:
                match = subset.loc[subset.index.date == cell_value]
                if isinstance(match, pd.DataFrame) and len(match) > 1:
                    combined_commessa = "; ".join(sorted(set(match["COMMESSA"])))
                    combined_ore = match["ORE"].sum()
                else:
                    combined_commessa = match["COMMESSA"].values[0]
                    combined_ore = match["ORE"].values[0]

                commessa_cell.value = combined_commessa
                ore_cell.value = combined_ore

    # Save back to the same file (in place)
    wb.save(strategie_file)


def main():
    root = tk.Tk()
    root.withdraw()

    # 1) Ask for the timesheet file
    timesheet_file = filedialog.askopenfilename(
        title="Select Timesheet Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not timesheet_file:
        messagebox.showerror("Error", "No timesheet file selected. Exiting.")
        return

    # 2) Ask for the mapping file
    mapping_file = filedialog.askopenfilename(
        title="Select Codice Commessa Mapping Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not mapping_file:
        messagebox.showerror("Error", "No mapping file selected. Exiting.")
        return

    # 3) Ask for the StrategieDigitali file
    strategie_file = filedialog.askopenfilename(
        title="Select StrategieDigitali_2025 Excel File (to update in place)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not strategie_file:
        messagebox.showerror("Error", "No StrategieDigitali file selected. Exiting.")
        return

    # Build the aggregated data from the timesheet
    df_agg = build_records_from_timesheet(timesheet_file, mapping_file)
    if df_agg.empty:
        messagebox.showerror("Error", "No valid timesheet data found. Exiting.")
        return

    # Update the StrategieDigitali file in place with openpyxl
    try:
        update_strategie_in_place(strategie_file, df_agg)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during update:\n{e}")
        return

    # Show a popup indicating success
    messagebox.showinfo("Success", "Update complete. Your StrategieDigitali file has been updated.")

if __name__ == "__main__":
    main()
