import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import timedelta

def main():
    # Set up the Tkinter root and hide the main window
    root = tk.Tk()
    root.withdraw()

    # Prompt the user to select the input Excel file
    input_file = filedialog.askopenfilename(
        title="Select Input Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not input_file:
        print("No input file selected. Exiting...")
        return

    # Prompt the user to select where to save the output Excel file
    output_file = filedialog.asksaveasfilename(
        title="Save Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not output_file:
        print("No output file selected. Exiting...")
        return

    # -------------------------------------------------------------------
    # STEP A: Load the selected Excel file (the "Timesheet Attivi")
    # -------------------------------------------------------------------
    df_raw = pd.read_excel(input_file)

    # -------------------------------------------------------------------
    # STEP B: Set up a weekday-offset map (Italian weekdays)
    # -------------------------------------------------------------------
    day_offset = {
        "Lunedì": 0,
        "Martedì": 1,
        "Mercoledì": 2,
        "Giovedì": 3,
        "Venerdì": 4,
        "Sabato": 5,
        "Domenica": 6
    }

    # -------------------------------------------------------------------
    # STEP C: Map "Codice Commessa" to a shorter or different label if needed
    # -------------------------------------------------------------------
    commessa_map = {
        "I112 - SYS - SA/RC": "23WP030 Sa-Rc"
        # Add more mappings as needed...
    }

    # -------------------------------------------------------------------
    # STEP D: Convert the weekly timesheet into day-by-day records
    # -------------------------------------------------------------------
    records = []
    for idx, row in df_raw.iterrows():
        # 1) Grab the text "WeekRange" (e.g. "03/03/2025 al 09/03/2025")
        week_range_str = str(row.get("WeekRange", "")).strip()
        if " al " not in week_range_str:
            continue
        
        # 2) Parse the start date from the text
        try:
            start_str, _ = week_range_str.split(" al ")
            start_date = pd.to_datetime(start_str, dayfirst=True)
        except Exception as e:
            print(f"Skipping row {idx} due to date parse error: {e}")
            continue

        # 3) Get the "Codice Commessa" and map it if needed
        codice_commessa = row.get("Codice Commessa", "")
        commessa_final = commessa_map.get(codice_commessa, codice_commessa)

        # 4) Determine the "Autore" and extract the surname
        autore = str(row.get("Autore", "")).strip()
        surname = autore.split()[-1] if autore else "UNKNOWN"

        # 5) Process each weekday column for hours
        for day_col, offset in day_offset.items():
            hours = row.get(day_col, 0)
            if pd.notna(hours) and float(hours) != 0.0:
                actual_date = start_date + timedelta(days=offset)
                records.append({
                    "DATA": actual_date.date(),
                    "COMMESSA": commessa_final,
                    "ORE": float(hours),
                    "SURNAME": surname
                })

    # -------------------------------------------------------------------
    # STEP E: Create a DataFrame from records and group duplicate entries
    # -------------------------------------------------------------------
    df_days = pd.DataFrame(records)
    df_final = df_days.groupby(["DATA", "COMMESSA", "SURNAME"], as_index=False).sum("ORE")

    # -------------------------------------------------------------------
    # STEP F: Write the output Excel file with one sheet per surname
    # -------------------------------------------------------------------
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for author_surname, sub_df in df_final.groupby("SURNAME"):
            sub_df_to_save = sub_df.drop(columns=["SURNAME"])
            sheet_name = author_surname[:31]  # Excel sheet name max length is 31
            sub_df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Created '{output_file}' with one sheet per surname.")

if __name__ == "__main__":
    main()
