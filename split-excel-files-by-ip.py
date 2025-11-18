import os
import pandas as pd

input_file_path = "C:/Users/DELL/Documents/DataFi/split documents/doc/Comined_FY25Q3_flatfile_CS.xlsx"  #NDR_linelist/usaid_patient_biometric_dsd_26_8_2025.csv"
output_folder_name = "C:/Users/DELL/Documents/DataFi/split documents/splitted_CS"  #splitted_files"
# column to split the data by
split_column_name = 'IP'

# Creates the output folder if it doesn't exist
if not os.path.exists(output_folder_name):
    os.makedirs(output_folder_name)
    print(f"Created folder: {output_folder_name}")

try:
    print(f"Reading file: {input_file_path}...")
    
    df = pd.read_excel(input_file_path)#, encoding='latin1', engine='python', on_bad_lines='skip')
    print("File read successfully.")

    # Group the DataFrame by the specified column
    groups = df.groupby(split_column_name)
    print(f"Found {len(groups)} unique values in the '{split_column_name}' column.")

    # Iterate over each group and save it to a separate file
    for group_name, group_df in groups:
        
        safe_group_name = str(group_name).replace(':', '_').replace('.', '_').replace('/', '_')
        
        # output file name
        output_file_name = f"{safe_group_name}_data.xlsx"
        output_file_path = os.path.join(output_folder_name, output_file_name)
        
        # Save to a new Excel file
        group_df.to_excel(output_file_path, index=False)
        
        print(f"Saved {output_file_name} with {len(group_df)} rows.")

    print("All files have been successfully split by IP.")

except FileNotFoundError:
    print(f"Error: The file '{input_file_path}' was not found.")
except KeyError:
    print(f"Error: The column '{split_column_name}' was not found in the CSV file.")
except Exception as e:
    print(f"An error occurred: {e}")