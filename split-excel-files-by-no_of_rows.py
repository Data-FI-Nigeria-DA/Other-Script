import os
import pandas as pd


input_file_path = "C:/Users/DELL/Documents/DataFi/NDR_linelist/usaid_patient_biometric_dsd_26_8_2025.csv"
rows_per_file = 500000
output_folder_name = "C:/Users/DELL/Documents/DataFi/splitted_files"

# Creates the output folder if it doesn't exist
if not os.path.exists(output_folder_name):
    os.makedirs(output_folder_name)
    print(f"Created folder: {output_folder_name}")

try:
    print(f"Reading file in chunks: {input_file_path}...")
    
    # chunksize to read the file in smaller parts
    chunk_reader = pd.read_csv(
        input_file_path,
        chunksize=rows_per_file,  # The size of each chunk
        encoding='latin1',
        engine='python',
        on_bad_lines='skip'
    )
    
    file_count = 0
    # Iterate through the chunks
    for i, chunk_df in enumerate(chunk_reader):
        file_count += 1
        
        # output file name
        output_file_name = f"part_{file_count}.xlsx"
        output_file_path = os.path.join(output_folder_name, output_file_name)
        
        # Save the chunk to a new Excel file
        chunk_df.to_excel(output_file_path, index=False)
        
        print(f"Saved {output_file_name} with {len(chunk_df)} rows.")
        
    print("All files have been successfully split and saved.")
    
except FileNotFoundError:
    print(f"Error: The file '{input_file_path}' was not found.")
except Exception as e:
    print(f"An error occurred: {e}")