import pandas as pd

# === 1. Load the Excel File ===
file_path = "/Users/dhairyasinghal/Desktop/Project/Exec Ed/TECK Dhairya Sheet.xlsx"
df = pd.read_excel(file_path, sheet_name='Students')

# === 2. Select Key Columns ===
columns_to_keep = [
    "Student Number", "First Name", "Last Name", "Full Name", "Operation",
    "GDBA?", "GDBA Admit Term", "GDBA Admit Date", "GDBA Completed Term",
    "GDBA Complete Date", "GDBA Duration (Years)",
    "EMBA?", "EBMA Admit Term", "EMBA Admit Date", "EMBA Completed Term",
    "EMBA Complete Date", "EMBA Duration (Years)", "Degree Awarded", "TECK or EVR"
]

df_clean = df[columns_to_keep].copy()

# === 3. Rename Columns ===
df_clean.columns = [
    "Student Number", "First Name", "Last Name", "Full Name", "Operation",
    "In GDBA", "GDBA Admit Term", "GDBA Admit Date", "GDBA Completed Term",
    "GDBA Complete Date", "GDBA Duration (Years)",
    "In EMBA", "EMBA Admit Term", "EMBA Admit Date", "EMBA Completed Term",
    "EMBA Complete Date", "EMBA Duration (Years)", "Degree Awarded", "Employer"
]

# === 4. Add Program Status Column ===
def determine_status(row):
    if row["In GDBA"] == "Y" and row["In EMBA"] == "Y":
        return "GDBA + EMBA"
    elif row["In GDBA"] == "Y":
        return "GDBA Only"
    elif row["In EMBA"] == "Y":
        return "EMBA Only"
    else:
        return "Unknown"

df_clean["Program Status"] = df_clean.apply(determine_status, axis=1)

# === 5. Create Subsets ===
gdbas = df_clean[df_clean["In GDBA"] == "Y"].copy()
embas = df_clean[df_clean["In EMBA"] == "Y"].copy()

# === 6. Save All Files ===
df_clean.to_excel("TECK_Master_Students.xlsx", index=False)
gdbas.to_excel("TECK_GDBA_Only.xlsx", index=False)
embas.to_excel("TECK_EMBA_Only.xlsx", index=False)

print("✅ All files exported successfully:")
print(" - TECK_Master_Students.xlsx")
print(" - TECK_GDBA_Only.xlsx")
print(" - TECK_EMBA_Only.xlsx")
