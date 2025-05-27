import pandas as pd

# === STEP 1: Load the Excel file ===
file_path = "/Users/dhairyasinghal/Desktop/Project/Exec Ed/Past Mitacs i2i Students Dhairya Sheet.xlsx"
xls = pd.ExcelFile(file_path)

# === STEP 2: Load both sheets ===
alumni_df = xls.parse("Alumni")
cohort_df = xls.parse("2022 i2i Cohort.")

# === STEP 3: Clean column names ===
alumni_df.columns = [col.strip() for col in alumni_df.columns]
cohort_df.columns = [col.strip() for col in cohort_df.columns]

# === STEP 4: Extract columns from Alumni ===
alumni_subset = alumni_df[[
    "First Name", "Last Name", "Contact Email", "Graduation Year i2I", "Job Title", "Company"
]].copy()
alumni_subset.rename(columns={
    "Contact Email": "Email",
    "Graduation Year i2I": "Cohort Year"
}, inplace=True)

# === STEP 5: Extract columns from 2022 Cohort ===
cohort_subset = cohort_df[[
    "First Name", "Last Name", "Contact Email", "Job Title", "Company"
]].copy()
cohort_subset.rename(columns={"Contact Email": "Email"}, inplace=True)
cohort_subset["Cohort Year"] = "2022"

# === STEP 6: Combine both sheets ===
combined_df = pd.concat([alumni_subset, cohort_subset], ignore_index=True)

# === STEP 7: Export to Excel ===
output_file = "Mitacs_i2i_Combined.xlsx"
combined_df.to_excel(output_file, index=False)

print(f"✅ File saved successfully: {output_file}")
