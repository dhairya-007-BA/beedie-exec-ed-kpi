import pandas as pd

# === STEP 1: Load Excel ===
file_path = "/Users/dhairyasinghal/Desktop/Project/file/Credential List Dhairya Sheet.xlsx"
xls = pd.ExcelFile(file_path)

# === STEP 2: Select valid course sheets ===
valid_sheets = [
    s for s in xls.sheet_names
    if s.startswith("Course") and "Removed" not in s and "Deferred" not in s and "Available" not in s
]

# === STEP 3: Extract the required columns ===
combined_data = []

for sheet in valid_sheets:
    try:
        df = xls.parse(sheet, header=1)
        df.columns = [str(col).strip() for col in df.columns]

        # Match expected columns (case-insensitive match)
        name_col = next((c for c in df.columns if "name" in c.lower()), None)
        email_col = next((c for c in df.columns if "email" in c.lower()), None)
        region_col = next((c for c in df.columns if "region" in c.lower()), None)
        service_col = next((c for c in df.columns if "service" in c.lower()), None)
        cohort_col = next((c for c in df.columns if "cohort" in c.lower()), None)

        if name_col and email_col and region_col and service_col and cohort_col:
            subset = df[[name_col, email_col, region_col, service_col, cohort_col]].copy()
            subset.columns = ["Name", "Email", "Region", "Service Line", "Cohort"]

            # Extract Track from sheet name (e.g., Course 2 - BES)
            parts = sheet.split(" -")
            track = parts[1].strip() if len(parts) > 1 else "Unknown"
            subset["Track"] = track

            combined_data.append(subset)

    except Exception as e:
        print(f"⚠️ Skipped {sheet}: {e}")

# === STEP 4: Combine and Export ===
final_df = pd.concat(combined_data, ignore_index=True)
output_file = "All_Courses_With_Cohort_Column.xlsx"
final_df.to_excel(output_file, index=False)

print(f"✅ File saved: {output_file}")
