import os
import pandas as pd
import re

# TO JOIN THE ALL_LD AND ALL_TAR FILES:

# Define the paths to the folders
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1"

# Create the output folder if it doesn't exist
output_folder = os.path.join(folder_path, "LD_with_TAR")
os.makedirs(output_folder, exist_ok=True)

# Define the paths to the Excel files
tars_file_path = r"C:\Users\nroselza\Downloads\LDDS_1\Provided files\TARs\ALL approved TAR.xlsx"
extractables_file_path = r"C:\Users\nroselza\Downloads\LDDS_1\Provided files\ALL_tests\LD_All Extractables.xlsx"
pyro_file_path = r"C:\Users\nroselza\Downloads\LDDS_1\Provided files\ALL_tests\LD_All PYRO.xlsx"
tdgcms_file_path = r"C:\Users\nroselza\Downloads\LDDS_1\Provided files\ALL_tests\LD_All TDGCMS.xlsx"
voc_file_path = r"C:\Users\nroselza\Downloads\LDDS_1\Provided files\ALL_tests\LD_All VOC.xlsx"

# Load the Excel files into pandas DataFrames
tars_df = pd.read_excel(tars_file_path, engine='openpyxl')
extractables_df = pd.read_excel(extractables_file_path, engine='openpyxl')
pyro_df = pd.read_excel(pyro_file_path, engine='openpyxl')
tdgcms_df = pd.read_excel(tdgcms_file_path, engine='openpyxl')
voc_df = pd.read_excel(voc_file_path, engine='openpyxl')

# List of DataFrames to modify
all_tests_dfs = {
    "extractables_df": extractables_df,
    "pyro_df": pyro_df,
    "tdgcms_df": tdgcms_df,
    "voc_df": voc_df
}

# Function to iterate through all rows of 'Item Number (Affected Items)' in tars_df with rows of a given DataFrame
def iterate_and_modify(tars_df, target_df, target_name):
    for tars_row_index in range(len(tars_df)):
        tars_row = tars_df.iloc[tars_row_index]
        item_number_affected = str(tars_row['Item Number (Affected Items)'])
        
        # Print the string value from the current row being iterated in tars_df
        print(f"Iterating TARs row {tars_row_index}: {item_number_affected}")
        
        # Compile the regex pattern for the search string
        pattern = re.escape(item_number_affected)
        
        # Iterate through all rows of target_df and search for the string in 'Item Description (Items)' and 'Item Number (Items)'
        mask = target_df['Item Description (Items)'].astype(str).str.contains(pattern, na=False, regex=True) | \
               target_df['Item Number (Items)'].astype(str).str.contains(pattern, na=False, regex=True)
        
        # Add the columns from TARs to the matching rows in target_df
        for col in tars_df.columns:
            if col not in target_df.columns:
                target_df[col] = None
            target_df.loc[mask, col] = tars_row[col]
    
    # Save the modified target_df to the output folder
    output_file_path = os.path.join(output_folder, f"LD_All {target_name}.xlsx")
    target_df.to_excel(output_file_path, index=False)
    
    print(f"Modified LD_All {target_name}.xlsx saved to {output_folder}")

# Iterate and modify extractables_df, pyro_df, and tdgcms_df with tars_df
iterate_and_modify(tars_df, extractables_df, "Extractables")
iterate_and_modify(tars_df, pyro_df, "PYRO")
iterate_and_modify(tars_df, tdgcms_df, "TDGCMS")

# SAME ONLY FOR VOC

# Iterate through all rows of 'Item Number (Affected Items)' in the TARs file with rows of voc_df
for tars_row_index in range(len(tars_df)):
    tars_row = tars_df.iloc[tars_row_index]
    item_number_affected = str(tars_row['Item Number (Affected Items)'])
    
    # Print the string value from the current row being iterated in tars_df
    print(f"Iterating TARs row {tars_row_index}: {item_number_affected}")
    
    # Compile the regex pattern for the search string
    pattern = re.escape(item_number_affected)
    
    # Iterate through all rows of voc_df and search for the string in 'Item Description (Items)' and 'Item Number (Items)'
    mask = voc_df['Item Description (Items)'].astype(str).str.contains(pattern, na=False, regex=True) | \
           voc_df['Item Number (Items)'].astype(str).str.contains(pattern, na=False, regex=True)
    
    # Add the columns from TARs to the matching rows in voc_df
    for col in tars_df.columns:
        if col not in voc_df.columns:
            voc_df[col] = None
        voc_df.loc[mask, col] = tars_row[col]

# Save the modified voc_df to the output folder
output_file_path = os.path.join(output_folder, "LD_All VOC.xlsx")
voc_df.to_excel(output_file_path, index=False)

print(f"Modified LD_All VOC.xlsx saved to {output_folder}")


# TO JOIN THE ALREADY JOINT FILES WITH THE SUBSTANCES' CAS NUMBERS:

# Path to the folder with modified files
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1"
ld_with_tar_folder = os.path.join(folder_path, "LD_with_TAR")

# Create the output folder if it doesn't exist
final_folder = os.path.join(folder_path, "Final")
os.makedirs(final_folder, exist_ok=True)

# Load the modified Excel files into pandas DataFrames
extractables_df = pd.read_excel(os.path.join(ld_with_tar_folder, "LD_All Extractables.xlsx"), engine='openpyxl')
pyro_df = pd.read_excel(os.path.join(ld_with_tar_folder, "LD_All PYRO.xlsx"), engine='openpyxl')
tdgcms_df = pd.read_excel(os.path.join(ld_with_tar_folder, "LD_All TDGCMS.xlsx"), engine='openpyxl')
voc_df = pd.read_excel(os.path.join(ld_with_tar_folder, "LD_All VOC.xlsx"), engine='openpyxl')

# List of DataFrames to modify
all_tests_dfs = {
    "extractables_df": extractables_df,
    "pyro_df": pyro_df,
    "tdgcms_df": tdgcms_df,
    "voc_df": voc_df
}

# Path to the folder with provided original files
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1\Provided files"

# Function to iterate through all rows of 'Name' in the target DataFrame and add matching rows from P1 and CC files
def add_matching_rows(target_df, target_name):
    # Add new columns for Description_LD, Status_LD, and Material type if not exists
    if 'Description_LD' not in target_df.columns:
        target_df['Description_LD'] = None
    if 'Status_LD' not in target_df.columns:
        target_df['Status_LD'] = None
    if 'Material type' not in target_df.columns:
        target_df['Material type'] = None

    # Iterate through each subfolder (P1 and CC)
    for subfolder in ["P1", "CC"]:
        # Get the full path of each subfolder
        subfolder_path = os.path.join(folder_path, subfolder)
        
        # Check if the subfolder exists
        if os.path.exists(subfolder_path):
            # Iterate through each file in the subfolder
            for file_name in os.listdir(subfolder_path):
                # Check if the file is an Excel file
                if file_name.endswith(".xlsx"):
                    # Create the full path of each Excel file
                    file_path = os.path.join(subfolder_path, file_name)
                    
                    # Load the Excel file into a pandas DataFrame
                    df = pd.read_excel(file_path, engine='openpyxl')
                    
                    # Ensure column names are stripped of leading/trailing spaces
                    df.columns = df.columns.str.strip()
                    
                    # Print column names to verify 'Name' exists
                    print(f"Columns in {file_name}: {df.columns.tolist()}")
                    
                    # Iterate through each row in the target DataFrame
                    for index, target_row in target_df.iterrows():
                        name_value = str(target_row['Name']).strip()
                        
                        # Print the line it is iterating
                        print(f"Iterating {target_name} row {index}: {name_value}")
                        
                        # Search for the value in the first column (Name) of the current DataFrame (P1 or CC)
                        mask = df['Name'].astype(str).str.strip() == name_value
                        
                        # If a match is found, add the rows from P1 or CC to the row of the target DataFrame in question
                        if mask.any():
                            for col in df.columns:
                                if col == 'Name':
                                    continue  # Skip adding 'Name' column again
                                elif col == 'Description':
                                    target_df.loc[index, 'Description_LD'] = df.loc[mask, col].values[0]
                                elif col == 'Status':
                                    target_df.loc[index, 'Status_LD'] = df.loc[mask, col].values[0]
                                else:
                                    if col not in target_df.columns:
                                        target_df[col] = None
                                    target_df.loc[index, col] = df.loc[mask, col].values[0]
                            # Add the name of the file (from CC or P1) to the Material type column
                            target_df.loc[index, 'Material type'] = file_name

    # Save the final modified DataFrame to the output folder
    target_df.to_excel(os.path.join(final_folder, f"Final_LD_All {target_name}.xlsx"), index=False)
    print(f"Final modified {target_name} file saved to {final_folder}")

# Iterate and modify extractables_df, tdgcms_df, and voc_df with matching rows from P1 and CC files
add_matching_rows(extractables_df, "Extractables")
add_matching_rows(tdgcms_df, "TDGCMS")
add_matching_rows(voc_df, "VOC")

# SAME ONLY FOR PYRO:

# Add new columns for Description_LD, Status_LD, and Material type if not exists
if 'Description_LD' not in pyro_df.columns:
    pyro_df['Description_LD'] = None
if 'Status_LD' not in pyro_df.columns:
    pyro_df['Status_LD'] = None
if 'Material type' not in pyro_df.columns:
    pyro_df['Material type'] = None

# Iterate through each subfolder (P1 and CC)
for subfolder in ["P1", "CC"]:
    # Get the full path of each subfolder
    subfolder_path = os.path.join(folder_path, subfolder)
    
    # Check if the subfolder exists
    if os.path.exists(subfolder_path):
        # Iterate through each file in the subfolder
        for file_name in os.listdir(subfolder_path):
            # Check if the file is an Excel file
            if file_name.endswith(".xlsx"):
                # Create the full path of each Excel file
                file_path = os.path.join(subfolder_path, file_name)
                
                # Load the Excel file into a pandas DataFrame
                df = pd.read_excel(file_path, engine='openpyxl')
                
                # Ensure column names are stripped of leading/trailing spaces
                df.columns = df.columns.str.strip()
                
                # Print column names to verify 'Name' exists
                print(f"Columns in {file_name}: {df.columns.tolist()}")
                
                # Iterate through each row in the ALL PYRO DataFrame
                for index, pyro_row in pyro_df.iterrows():
                    name_value = str(pyro_row['Name']).strip()
                    
                    # Print the line it is iterating
                    print(f"Iterating ALL PYRO row {index}: {name_value}")
                    
                    # Search for the value in the first column (Name) of the current DataFrame (P1 or CC)
                    mask = df['Name'].astype(str).str.strip() == name_value
                    
                    # If a match is found, add the rows from P1 or CC to the row of ALL PYRO in question
                    if mask.any():
                        for col in df.columns:
                            if col == 'Name':
                                continue  # Skip adding 'Name' column again
                            elif col == 'Description':
                                pyro_df.loc[index, 'Description_LD'] = df.loc[mask, col].values[0]
                            elif col == 'Status':
                                pyro_df.loc[index, 'Status_LD'] = df.loc[mask, col].values[0]
                            else:
                                if col not in pyro_df.columns:
                                    pyro_df[col] = None
                                pyro_df.loc[index, col] = df.loc[mask, col].values[0]
                        # Add the name of the file (from CC or P1) to the Material type column
                        pyro_df.loc[index, 'Material type'] = file_name

# Save the final modified ALL PYRO DataFrame to the output folder
pyro_df.to_excel(os.path.join(final_folder, "Final_LD_All PYRO.xlsx"), index=False)

print(f"Final modified ALL PYRO file saved to {final_folder}")

# TO CHANGE COLUMNS' POSITIONS:

# Path to the folder with provided files
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1\Final"

# Create the output folder if it doesn't exist
final_columns_folder = os.path.join(folder_path, "Final with final columns")
os.makedirs(final_columns_folder, exist_ok=True)

# List of files to process
files_to_process = ["Final_LD_All Extractables.xlsx", "Final_LD_All VOC.xlsx", "Final_LD_All TDGCMS.xlsx", "Final_LD_All PYRO.xlsx"]

# Desired column order
desired_columns = [
    "CAS Number (Item Composition)", "Material type", "Name", "Status_LD", "Number", "Status",
    "Item Number (Items)", "Item Description (Items)", "Item Number (Affected Items)", "Item Description (Affected Items)",
    "Description", "Description_LD", "Calculated PPM (Item Composition)", "Comments (Lab Discovery Details)",
    "Use Classification (Request Details)", "PI2K Use Classification (Request Details)", "Vendor or Laboratory",
    "Workflow", "Compliance Manager", "Declaration Type"
]

# Process each file
for file_name in files_to_process:
    # Load the Excel file into a pandas DataFrame
    file_path = os.path.join(folder_path, file_name)
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Reorder the columns according to the desired order
    df = df.reindex(columns=desired_columns)
    
    # Save the modified DataFrame to the output folder
    output_file_path = os.path.join(final_columns_folder, file_name)
    df.to_excel(output_file_path, index=False)

print(f"Files with final columns saved to {final_columns_folder}")

# TO FILTER OUT THE NON INTERESTING ROWS, FILTER 1:

# Path to the folder with provided files
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1\Final\Final with final columns"

# Create the output folder if it doesn't exist
real_final_folder = os.path.join(folder_path, "Real final")
os.makedirs(real_final_folder, exist_ok=True)

# List of files to process
files_to_process = ["Final_LD_All Extractables.xlsx", "Final_LD_All VOC.xlsx", "Final_LD_All TDGCMS.xlsx", "Final_LD_All PYRO.xlsx"]

# Columns to clear if Item Number (Affected Items) is empty but Number is not
columns_to_clear = [
    "Number", "Description", "Use Classification (Request Details)", "PI2K Use Classification (Request Details)",
    "Status", "Item Description (Affected Items)", "Item Number (Affected Items)"
]

# Process each file
for file_name in files_to_process:
    # Load the Excel file into a pandas DataFrame
    file_path = os.path.join(folder_path, file_name)
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Delete rows where CAS Number (Item Composition) is empty
    df = df.dropna(subset=["CAS Number (Item Composition)"])
    
    # Delete rows where CAS Number (Item Composition) is not empty, but Status_LD and Item Number (Affected Items) are both empty
    df = df[~((df["CAS Number (Item Composition)"].notna()) & (df["Status_LD"].isna()) & (df["Item Number (Affected Items)"].isna()))]
    
    # Clear specific columns if Item Number (Affected Items) is empty but Number is not
    mask = df["Item Number (Affected Items)"].isna() & df["Number"].notna()
    df.loc[mask, columns_to_clear] = None
    
    # Save the modified DataFrame to the output folder
    output_file_path = os.path.join(real_final_folder, file_name)
    df.to_excel(output_file_path, index=False)

print(f"Files with final modifications saved to {real_final_folder}")

# FILTER 2:

# Path to the folder with provided files
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1\Final\Final with final columns"

# Create the output folder if it doesn't exist
real_final_filter_2_folder = os.path.join(folder_path, "Real final_filter 2")
os.makedirs(real_final_filter_2_folder, exist_ok=True)

# List of files to process
files_to_process = ["Final_LD_All Extractables.xlsx", "Final_LD_All VOC.xlsx", "Final_LD_All TDGCMS.xlsx", "Final_LD_All PYRO.xlsx"]

# Columns to clear if Item Number (Affected Items) is empty but Number is not
columns_to_clear = [
    "Number", "Description", "Use Classification (Request Details)", "PI2K Use Classification (Request Details)",
    "Status", "Item Description (Affected Items)", "Item Number (Affected Items)"
]

# Process each file
for file_name in files_to_process:
    # Load the Excel file into a pandas DataFrame
    file_path = os.path.join(folder_path, file_name)
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Delete rows where Calculated PPM (Item Composition) has a value but CAS Number (Item Composition) is empty
    df = df[~((df["Calculated PPM (Item Composition)"].notna()) & (df["CAS Number (Item Composition)"].isna()))]
    
    # Write "No peak detected" in CAS Number (Item Composition) where it is empty
    df["CAS Number (Item Composition)"].fillna("No peak detected", inplace=True)
    
    # Clear specific columns if Item Number (Affected Items) is empty but Number is not
    mask = df["Item Number (Affected Items)"].isna() & df["Number"].notna()
    df.loc[mask, columns_to_clear] = None
    
    # Save the modified DataFrame to the output folder
    output_file_path = os.path.join(real_final_filter_2_folder, file_name)
    df.to_excel(output_file_path, index=False)

print(f"Files with final modifications saved to {real_final_filter_2_folder}")

# TO MERGE THE FOUR FILES:

# Path to the folder with provided files
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1\Final\Final with final columns\Real final_filter 2"

# List of files to process
files_to_process = [
    ("Final_LD_All Extractables.xlsx", "Extractables"),
    ("Final_LD_All VOC.xlsx", "VOC"),
    ("Final_LD_All TDGCMS.xlsx", "TDGCMS"),
    ("Final_LD_All PYRO.xlsx", "PYRO")
]

# Initialize an empty DataFrame to store the merged data
merged_df = pd.DataFrame()

# Process each file and add a column for Test type
for file_name, test_type in files_to_process:
    # Load the Excel file into a pandas DataFrame
    file_path = os.path.join(folder_path, file_name)
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Add a column for Test type
    df['Test type'] = test_type
    
    # Append the DataFrame to the merged DataFrame
    merged_df = pd.concat([merged_df, df], ignore_index=True)

# Save the merged DataFrame to the output folder
output_file_path = os.path.join(folder_path, "MERGED_ALL_LDs.xlsx")
merged_df.to_excel(output_file_path, index=False)

print(f"Merged file saved to {output_file_path}")

# TO ADD THE TARs FROM THE COMPLEMENTARY TARs FILE

# Path to the folder with provided files
folder_path = r"C:\Users\nroselza\Downloads\LDDS_1\Provided files\TARs"

# Load the MERGED_ALL_LDs file
merged_file_path = r"C:\Users\nroselza\Downloads\LDDS_1\Final\Final with final columns\Real final_filter 2\MERGED_ALL_LDs.xlsx"
if not os.path.exists(merged_file_path):
    print(f"File not found: {merged_file_path}")
else:
    merged_df = pd.read_excel(merged_file_path, engine='openpyxl')

# Load the 20250124_Complementary TAR list file
complementary_tar_file_path = os.path.join(folder_path, "20250124_Complementary TAR list.xlsx")
complementary_tar_df = pd.read_excel(complementary_tar_file_path, engine='openpyxl')

# Print column names to verify
print(complementary_tar_df.columns)

# Iterate through each row in the complementary TAR DataFrame
for index, row in complementary_tar_df.iterrows():
    value_b = row['B'].strip()  # Column B in complementary TAR, strip trailing spaces
    
    # Search for the value in the MERGED_ALL_LDs DataFrame's column C (named "Name") and column G (named "Item Number (Items)")
    mask_name = merged_df['Name'] == value_b
    mask_item_number = merged_df['Item Number (Items)'] == value_b
    
    # If a match is found in either column, add the values from complementary TAR to the MERGED_ALL_LDs DataFrame
    if mask_name.any() or mask_item_number.any():
        merged_df.loc[mask_name | mask_item_number, 'Number'] = row['A']  # Column A to Number (column E)
        merged_df.loc[mask_name | mask_item_number, 'Status_LD'] = row['F']  # Column F to Status_LD (column D)
        merged_df.loc[mask_name | mask_item_number, 'Declaration Type'] = row['J']  # Column J to Declaration Type (column T)

# Save the modified DataFrame to the output folder
output_file_path = r"C:\Users\nroselza\Downloads\LDDS_1\Final\Final with final columns\Real final_filter 2\FINAL_MERGED_ALL_LDs.xlsx"
merged_df.to_excel(output_file_path, index=False)

print(f"Final merged file saved to {output_file_path}")