import pandas as pd

# === STEP 1: Load files ===
file1 = "/Users/dhairyasinghal/Desktop/Project/Exec Ed/TLP 15- All Applications.xlsx"
file2 = "/Users/dhairyasinghal/Desktop/Project/Exec Ed/TLP Program Enrollment Status Updates.xls.xlsx"

xls1 = pd.ExcelFile(file1)
xls2 = pd.ExcelFile(file2)

# === STEP 2: Load sheets ===
df_completion = xls1.parse("Completion List")
df_applications = xls1.parse("TLP Current Applications")
df_industry = xls1.parse("Industry-Location")
df_enrollment = xls2.parse("TLP Program Enrollment Updates")

# === STEP 3: Clean column names ===
for df in [df_completion, df_applications, df_industry, df_enrollment]:
    df.columns = df.columns.str.strip()

# === STEP 4: Extract and standardize key columns from each ===

# Completion List
completion = df_completion[["Applicant Name", "Email", "Complete/Incomplete", "Title", "Company"]].copy()
completion.rename(columns={"Applicant Name": "Name"}, inplace=True)

# Current Applications
applications = df_applications[["Applicant Name", "SFU Review Comments"]].copy()
applications.rename(columns={"Applicant Name": "Name"}, inplace=True)

# Industry Location
industry = df_industry[["Applicant Name", "SFU Review Comments"]].copy()
industry.rename(columns={"Applicant Name": "Name", "SFU Review Comments": "SFU Review Comments Industry"}, inplace=True)

# Enrollment Updates
enrollment = df_enrollment[["Program Enrollment ID", "Contact", "Program", "Program Offering", "Status"]].copy()

# === STEP 5: Build a master list of all names ===
names_from_completion = completion["Name"]
names_from_applications = applications["Name"]
names_from_industry = industry["Name"]
names_from_enrollment = enrollment["Contact"]

all_names = pd.Series(pd.concat([
    names_from_completion, names_from_applications, names_from_industry, names_from_enrollment
]).dropna().unique(), name="Name")

master = pd.DataFrame(all_names)

# === STEP 6: Merge everything onto the master name list ===
merged = master.merge(completion, on="Name", how="left")
merged = merged.merge(applications, on="Name", how="left")
merged = merged.merge(industry, on="Name", how="left")
merged = merged.merge(enrollment, left_on="Name", right_on="Contact", how="left")

# === STEP 7: Build final output ===
final = merged[[
    "Program Enrollment ID",
    "Name",
    "Email",
    "Contact",
    "Title",
    "Company",
    "Complete/Incomplete",
    "Program",
    "Program Offering"
]].copy()

# Fill SFU Review Comments from either source
final["SFU Review Comments"] = merged["SFU Review Comments"].combine_first(merged["SFU Review Comments Industry"])

# === STEP 8: Export ===
output_file = "TLP_Master_Merged_Final.xlsx"
final.to_excel(output_file, index=False)

print(f"✅ Success! File saved as: {output_file}")
