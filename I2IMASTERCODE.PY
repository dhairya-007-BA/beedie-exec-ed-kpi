import pandas as pd
import re

# === STEP 1: Load Excel files ===
file1 = "/Users/dhairyasinghal/Desktop/Project/file/filename.xlsx"
file2 = "/Users/dhairyasinghal/Desktop/Project/file/filename.xlsx"
file3 = "/Users/dhairyasinghal/Desktop/Project/file/filename.xlsx"

xls1 = pd.ExcelFile(file1)
xls2 = pd.ExcelFile(file2)
xls3 = pd.ExcelFile(file3)

df1 = pd.read_excel(xls1, sheet_name="Full Student List")
df2 = pd.read_excel(xls2, sheet_name="file")
df3 = pd.read_excel(xls3, sheet_name="file")

# === STEP 2: Clean column names ===
df1.columns = [col.strip() for col in df1.columns]
df2.columns = [col.strip() for col in df2.columns]
df3.columns = [col.strip() for col in df3.columns]

# === STEP 3: Select and rename relevant columns ===
df1 = df1[["First Name", "Last Name", "Student #", "Email", "Year", "Status"]].copy()
df1.rename(columns={"Student #": "Student Number", "Cohort": "Year"}, inplace=True)

df2 = df2[["First Name", "Last Name", "SFU Email", "Alternate Emails", "Cohort"]].copy()
df2.rename(columns={
    "Alternate Emails": "Personal Email",
    "Cohort": "Cohort Year"
}, inplace=True)

df3 = df3[["firstName", "lastName", "email", "Status"]].copy()
df3.rename(columns={
    "firstName": "First Name",
    "lastName": "Last Name",
    "email": "SFU Email"
}, inplace=True)

# === STEP 4: Merge all datasets ===
merged = df1.merge(df2, on=["First Name", "Last Name", "Email"], how="outer")
merged = merged.merge(df3, on=["First Name", "Last Name", "Email"], how="outer")

# === STEP 5: Final cleanup ===
# Combine Status and Cohort Year from both sources
final = pd.DataFrame({
    "First Name": merged["First Name"],
    "Last Name": merged["Last Name"],
    "Student Number": merged.get("Student Number"),
    "SFU Email": merged["Email"],
    "Personal Email": merged.get("Personal Email"),
    "Cohort Year": merged["Cohort Year_x"].combine_first(merged.get("Cohort Year_y")),
    "Status": merged["Status_x"].combine_first(merged.get("Status_y"))
})

# === STEP 6: Remove duplicates ===
final.drop_duplicates(inplace=True)

# === STEP 7: Clean names (keep only first and last word) ===
final["First Name"] = final["First Name"].astype(str).str.strip().str.split().str[0]
final["Last Name"] = final["Last Name"].astype(str).str.strip().str.split().str[-1]

# === STEP 8: Clean emails (remove mailto:) ===
final["SFU Email"] = final["Email"].astype(str).str.replace("mailto:", "", regex=False).str.strip()
final["Personal Email"] = final["Personal Email"].astype(str).str.replace("mailto:", "", regex=False).str.strip()

# === STEP 9: Export to Excel ===
output_file = "Master_Cleaned.xlsx"
final.to_excel(output_file, index=False)
print(f"✅ File saved: {output_file}")
