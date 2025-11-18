import os
import pandas as pd

# Define paths
input_folder = "C:/Users/DELL/Documents/DataFi/splitted_files"  
output_parent_folder = "C:/Users/DELL/Documents/DataFi/Extracted_Data"

# List of files to process
files_to_process = ['ACE1_NDR.xlsx', 'ACE2_NDR.xlsx', 'ACE3_NDR.xlsx'] 

# columns to extract
columns_to_extract = ['Facility Name', 'Patient ID', 'Visit Date', 'Weight'] 

# parent output folder if it doesn't exist
if not os.path.exists(output_parent_folder):
    os.makedirs(output_parent_folder)
    print(f"Created parent folder: {output_parent_folder}")

# Loop through each file
for file_name in files_to_process:
    try:
        # input file path
        input_file_path = os.path.join(input_folder, file_name)

        
        file_base_name = os.path.splitext(file_name)[0]
        output_folder = os.path.join(output_parent_folder, file_base_name)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"Created output folder for {file_name}: {output_folder}")

        # Read the Excel file
        print(f"Reading file: {file_name}...")
        df = pd.read_excel(input_file_path)

        # desired columns
        extracted_df = df[columns_to_extract]

        # output file path
        output_file_path = os.path.join(output_folder, file_name)

        # Saves the new DataFrame to the new folder
        extracted_df.to_excel(output_file_path, index=False)
        
        print(f"Successfully extracted columns and saved to: {output_file_path}")

    except FileNotFoundError:
        print(f"Error: The file '{input_file_path}' was not found.")
    except KeyError as e:
        print(f"Error: The column {e} was not found in the file '{file_name}'. Skipping.")
    except Exception as e:
        print(f"An unexpected error occurred with file '{file_name}': {e}")

print("\nAll files processed.")