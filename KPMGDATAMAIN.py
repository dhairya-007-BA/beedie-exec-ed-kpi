import pandas as pd

# === STEP 1: Load the Excel file ===
file_path = "/Users/dhairyasinghal/Desktop/Project/Exec Ed/KPMG Masters and Certificate Dhairya Sheet.xlsx"
xls = pd.ExcelFile(file_path)

# === STEP 2: Load the 'Master Sheet' ===
df = xls.parse("Master Sheet")
df.columns = [col.strip() for col in df.columns]  # Clean whitespace from column names

# === STEP 3: Select required columns ===
columns_to_keep = [
    "First Name", "Last Name", "Student Number", "Cert Cohort Year", "KPMG Email",
    "SFU email", "KPMG Canada or China", "Student Status (SIMS)",
    "Certificate Completed (Y/N/ IP)", "Graduated with Certificate",
    "Term Certificate was Completed (e.g.,1237, 1247)",
    "MSc Cohort Year Starting (e.g., 2020, 22, 23,24)", "MSc Completed (Y/N/IP)",
    "Current KPMG Employee"
]

clean_df = df[columns_to_keep].copy()

# === STEP 4: Add Program Type column ===
def determine_program(row):
    cert = str(row["Certificate Completed (Y/N/ IP)"]).strip().upper()
    msc = str(row["MSc Completed (Y/N/IP)"]).strip().upper()
    if cert == 'Y' and msc == 'Y':
        return "Certificate + MSc"
    elif cert == 'Y':
        return "Certificate Only"
    elif msc == 'Y':
        return "MSc Only"
    else:
        return "None/Incomplete"

clean_df["Program Type"] = clean_df.apply(determine_program, axis=1)

# === STEP 5: Export cleaned data to Excel ===
output_file = "KPMG_Masters_Certificate_Analyzed.xlsx"
clean_df.to_excel(output_file, index=False)

print(f"✅ File exported successfully: {output_file}")
