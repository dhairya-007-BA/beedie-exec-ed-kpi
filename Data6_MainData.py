import pandas as pd

# === 1. Load the Excel File ===
file_path = "/Users/dhairyasinghal/Desktop/Project/file/filename.xlsx"
df = pd.read_excel(file_path, sheet_name='Students')

# === 2. Select Key Columns ===
columns_to_keep = [
    "Student Number", "First Name", "Last Name", "Full Name", "Operation",
    "File", "Admit Term", "Admit Date", "Completed Term",
    "Complete Date", "Duration (Years)",
    "file2", "Admit Term", "Admit Date", "Completed Term",
    "Complete Date", "Duration (Years)", "Degree Awarded", "file"
]

df_clean = df[columns_to_keep].copy()

# === 3. Rename Columns ===
df_clean.columns = [
    "Student Number", "First Name", "Last Name", "Full Name", "Operation",
    "In program", "Admit Term", "Admit Date", "Completed Term",
    "Complete Date", "Duration (Years)",
    "In program", "Admit Term", "Admit Date", "Completed Term",
    "Complete Date", "Duration (Years)", "Degree Awarded", "Employer"
]

# === 4. Add Program Status Column ===
def determine_status(row):
    if row["In program"] == "Y" and row["In program"] == "Y":
        return "program1 + program2"
    elif row["In program1"] == "Y":
        return "program1 Only"
    elif row["In program2"] == "Y":
        return "program2 Only"
    else:
        return "Unknown"

df_clean["Program Status"] = df_clean.apply(determine_status, axis=1)

# === 5. Create Subsets ===
gdbas = df_clean[df_clean["In program"] == "Y"].copy()
embas = df_clean[df_clean["In program"] == "Y"].copy()

# === 6. Save All Files ===
df_clean.to_excel("Master_Students.xlsx", index=False)


print("✅ All files exported successfully:")
print(" Master_Students.xlsx")
