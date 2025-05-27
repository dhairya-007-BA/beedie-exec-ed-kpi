import pandas as pd

# === STEP 1: Load the Excel file ===
file_path = "/Users/dhairyasinghal/Desktop/Project/file/filename.xlsx"
xls = pd.ExcelFile(file_path)

# === STEP 2: Load the 'Master Sheet' ===
df = xls.parse("Master Sheet")
df.columns = [col.strip() for col in df.columns]  # Clean whitespace from column names

# === STEP 3: Select required columns ===
columns_to_keep = [
    "First Name", "Last Name", "Student Number", "Year", "Email",
    "email", "Canada or China", "Student Status",
    "Certificate Completed (Y/N/ IP)", "Graduated with Cert",
    "Term Cert was Completed (e.g.,1237, 1247)",
    "Cohort Year Starting", "Completed (Y/N/IP)",
    "Current employee"
]

clean_df = df[columns_to_keep].copy()

# === STEP 4: Add Program Type column ===
def determine_program(row):
    cert = str(row["Cert Completed (Y/N/ IP)"]).strip().upper()
    msc = str(row["Completed (Y/N/IP)"]).strip().upper()
    if cert == 'Y' and msc == 'Y':
        return "Certificate + Masters"
    elif cert == 'Y':
        return "Certificate Only"
    elif msc == 'Y':
        return "MSc Only"
    else:
        return "None/Incomplete"

clean_df["Program Type"] = clean_df.apply(determine_program, axis=1)

# === STEP 5: Export cleaned data to Excel ===
output_file = "File_Analyzed.xlsx"
clean_df.to_excel(output_file, index=False)

print(f"✅ File exported successfully: {output_file}")
