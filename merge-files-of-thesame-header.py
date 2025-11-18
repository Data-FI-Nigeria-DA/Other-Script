import pandas as pd
import os

def combine_excel_files(input_folder, output_folder, output_filename="combined_data.xlsx"):
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # list of all Excel files in the input folder
    excel_files = [f for f in os.listdir(input_folder) if f.endswith(('.xls', '.xlsx', '.csv'))]

    if not excel_files:
        print("No Excel files found in the specified input folder.")
        return

    all_data = []

    # Loop through each file, read it into a DataFrame, and append to the list
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        if file_path.endswith('.csv'):
            df= pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        all_data.append(df)

    
    combined_df = pd.concat(all_data, ignore_index=True)

    # output path
    output_path = os.path.join(output_folder, output_filename)

    # Save the combined DataFrame to a new Excel file
    combined_df.to_excel(output_path, index=False)

    print(f"Successfully combined {len(excel_files)} files into '{output_path}'.")


if __name__ == "__main__":
    
    input_folder_path = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/Centralsync'
    output_folder_path = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis'

    combine_excel_files(input_folder_path, output_folder_path)