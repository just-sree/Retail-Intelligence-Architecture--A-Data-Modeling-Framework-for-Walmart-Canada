import os
import pandas as pd

# Parent folder containing sub-folders with CSV files
parent_folder_path = "D:\Personal Projects\Retail-Intelligence-Architecture--A-Data-Modeling-Framework-for-Walmart-Canada\Datasets"

# Iterate through each subfolder
for subfolder_name in os.listdir(parent_folder_path):
    subfolder_path = os.path.join(parent_folder_path, subfolder_name)
    
    # Check if it's a folder
    if os.path.isdir(subfolder_path):
        # Create an Excel file for the subfolder
        output_file = f"Datasets\{subfolder_name}.xlsx"
        
        # Create a Pandas ExcelWriter object
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Iterate through files in the subfolder
            for file_name in os.listdir(subfolder_path):
                if file_name.endswith('.csv'):  # Ensure the file is a CSV
                    # Read the CSV into a DataFrame
                    file_path = os.path.join(subfolder_path, file_name)
                    df = pd.read_csv(file_path)
                    
                    # Remove '.csv' from the file name for the sheet name
                    sheet_name = os.path.splitext(file_name)[0]
                    
                    # Write DataFrame to a sheet in the Excel file
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Processed all CSV files in '{subfolder_name}' and saved as '{output_file}'.")