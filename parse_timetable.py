"""
Step 1: One-time parser for KV Burhanpur timetable Excel file.
Produces timetable_master.csv with columns:
    Teacher_Name, Day, Period, Class, Subject
"""

import openpyxl
import pandas as pd

EXCEL_FILE = "Untitled spreadsheet.xlsx"
OUTPUT_CSV = "timetable_master.csv"

# 0-indexed column positions → period numbers
PERIOD_COLS = {2: 1, 3: 2, 5: 3, 6: 4, 8: 5, 9: 6, 10: 7, 11: 8}
DAYS = {"MON", "TUE", "WED", "THU", "FRI", "SAT"}


def clean(val):
    return str(val).strip() if val is not None else ""


def parse_timetable(excel_path: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    records = []
    current_teacher = None

    i = 0
    while i < len(rows):
        row = rows[i]

        # Teacher name appears in col[0] of the first day row of each block
        if row[0] is not None and clean(row[0]) != "" and row[1] in DAYS:
            current_teacher = clean(row[0])

        if row[1] in DAYS and current_teacher:
            day = row[1]
            class_row = row
            subj_row = rows[i + 1] if i + 1 < len(rows) else ("",) * 12

            for col_idx, period in PERIOD_COLS.items():
                cls = clean(class_row[col_idx])
                subj = clean(subj_row[col_idx])

                if cls or subj:
                    records.append(
                        {
                            "Teacher_Name": current_teacher,
                            "Day": day,
                            "Period": period,
                            "Class": cls,
                            "Subject": subj,
                        }
                    )
            i += 2  # skip the paired subject row
        else:
            i += 1

    df = pd.DataFrame(records, columns=["Teacher_Name", "Day", "Period", "Class", "Subject"])

    # Normalise whitespace in names
    df["Teacher_Name"] = df["Teacher_Name"].str.strip()
    df["Class"] = df["Class"].str.strip()
    df["Subject"] = df["Subject"].str.strip()

    return df


if __name__ == "__main__":
    df = parse_timetable(EXCEL_FILE)
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"Saved {len(df)} records to {OUTPUT_CSV}\n")
    print("Teachers found:")
    for t in sorted(df["Teacher_Name"].unique()):
        print(f"  {t}")
    print(f"\nSample rows:")
    print(df.head(20).to_string(index=False))
