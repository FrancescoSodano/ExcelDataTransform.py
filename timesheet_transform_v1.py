import pandas as pd
from datetime import timedelta

# -------------------------------------------------------------------
# STEP A: Load the raw Dataverse Excel (the "Timesheet Attivi")
# -------------------------------------------------------------------
timesheet_file = "Timesheets attivi-e 12-03-2025 11-38-25.xlsx"
df_raw = pd.read_excel(timesheet_file)

# -------------------------------------------------------------------
# STEP B: Set up a weekday-offset map (Italian weekdays)
#         "Lunedì" means 0 days from Monday, "Martedì" 1 day, etc.
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
# STEP C: If you need to map "Codice Commessa" to a shorter or different
#         label, do it here. Otherwise, you can skip or just keep them as is.
# -------------------------------------------------------------------
commessa_map = {
    "I112 - SYS - SA/RC": "23WP030 Sa-Rc"
    # Add more if needed...
}

# -------------------------------------------------------------------
# STEP D: Convert the weekly timesheet into day-by-day records
# -------------------------------------------------------------------
records = []

for idx, row in df_raw.iterrows():
    # 1) Grab the text "WeekRange" (e.g. "03/03/2025 al 09/03/2025")
    week_range_str = str(row.get("WeekRange", "")).strip()
    if " al " not in week_range_str:
        # Skip rows that don't match the "DD/MM/YYYY al DD/MM/YYYY" format
        continue
    
    # 2) Parse the start/end date from the text
    try:
        start_str, end_str = week_range_str.split(" al ")
        # "dayfirst=True" tells pandas that the format is D/M/YYYY
        start_date = pd.to_datetime(start_str, dayfirst=True)
        # end_date = pd.to_datetime(end_str,   dayfirst=True)  # If needed for checking
    except:
        continue  # If it fails to parse, skip

    # 3) Get the "Codice Commessa" and map it if needed
    codice_commessa = row.get("Codice Commessa", "")
    commessa_final = commessa_map.get(codice_commessa, codice_commessa)

    # 4) Figure out the "Autore" (the person). Example: "Pietro Fava"
    autore = row.get("Autore", "")
    autore = str(autore).strip()

    # We'll define the "sheet name" as the person's surname, 
    # i.e. the LAST "word" in "Pietro Fava" → "Fava"
    surname = autore.split()[-1] if autore else "UNKNOWN"

    # 5) For each weekday column, check if there's any hours
    for day_col, offset in day_offset.items():
        hours = row.get(day_col, 0)
        if pd.notna(hours) and float(hours) != 0.0:
            # The actual date is "start_date + offset days"
            actual_date = start_date + timedelta(days=offset)
            # Build a record for that day
            records.append({
                "DATA": actual_date.date(),   # or keep as a Timestamp
                "COMMESSA": commessa_final,
                "ORE": float(hours),
                "SURNAME": surname
            })

# -------------------------------------------------------------------
# STEP E: Convert the list of day-by-day dictionaries into a DataFrame
#         Then sum up any duplicates (same date, same commessa, same surname).
# -------------------------------------------------------------------
df_days = pd.DataFrame(records)
df_final = df_days.groupby(["DATA","COMMESSA","SURNAME"], as_index=False).sum("ORE")

# -------------------------------------------------------------------
# STEP F: We now have a day-by-day table with columns [DATA, COMMESSA, SURNAME, ORE].
#         We will write ONE sheet per "SURNAME" in a new Excel file.
# -------------------------------------------------------------------
output_file = "Timesheet_daybyday_perAuthor.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    # Group by surname so each person gets their own sheet
    for author_surname, sub_df in df_final.groupby("SURNAME"):
        # Drop the SURNAME column from the sub_df if you don't want it repeated
        # inside the sheet. Or keep it for clarity.
        sub_df_to_save = sub_df.drop(columns=["SURNAME"])
        # Let's name the sheet with the surname
        sheet_name = author_surname[:31]  # Excel sheet name max length = 31 chars
        sub_df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Created '{output_file}' with one sheet per surname.")

# -------------------------------------------------------------------
# STEP G: (Optional) If you want to "merge" these new day-by-day records into an
#         existing "StrategieDigitali_2025.xlsx" that also has multiple sheets
#         named "Fava", "Macchi", etc., you'll need a more custom approach:
#         1) open that file with openpyxl,
#         2) read each sheet with pandas,
#         3) merge or append sub_df_to_save,
#         4) overwrite the sheet.
#         This depends on how your "StrategieDigitali_2025.xlsx" is structured.
# -------------------------------------------------------------------
