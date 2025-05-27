import pandas as pd

# Load the Excel file containing 28 worksheets
excel_path = "/Users/dhairyasinghal/Desktop/Project/DEP Clean Data.xlsx"  # Update path if needed
excel_file = pd.ExcelFile(excel_path)

# Define the standard headers
expected_headers = ['First Name', 'Last Name', 'Email', 'Phone', 'Title', 'Company']

# Initialize an empty DataFrame with the correct columns
master_df = pd.DataFrame(columns=expected_headers)

# Iterate through all sheets and extract relevant columns if they exist
for sheet_name in excel_file.sheet_names:
    sheet_df = excel_file.parse(sheet_name)
    
    # Standardize the column names to match expected headers (case-insensitive match)
    renamed_columns = {}
    for col in sheet_df.columns:
        for expected in expected_headers:
            if col.strip().lower() == expected.lower():
                renamed_columns[col] = expected
                break
    
    # Rename and select only the columns that match expected headers
    sheet_df = sheet_df.rename(columns=renamed_columns)
    sheet_df = sheet_df[[col for col in expected_headers if col in sheet_df.columns]]

    # Add missing columns with empty values
    for col in expected_headers:
        if col not in sheet_df.columns:
            sheet_df[col] = None

    # Append to master DataFrame
    master_df = pd.concat([master_df, sheet_df[expected_headers]], ignore_index=True)

# Drop duplicate rows based on the primary key: 'First Name'
master_df = master_df.drop_duplicates(subset='First Name', keep='first')

# Optional: save the merged data to a new Excel file
master_df.to_excel("Merged_Master_Sheet.xlsx", index=False)
