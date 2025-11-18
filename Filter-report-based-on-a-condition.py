import pandas as pd
import os
from datetime import datetime
import re

'''Field to change are; folder_path, output_base_dir, Periods if necessary and Filter(condition for filtering)'''

# Path to the directory containing the CSV files
folder_path = 'C:/Users/DELL/Documents/DataFi/Data Review Meeting/Lab report Lookup to Radet/Lab report'

# Define output directory for projects
output_base_dir = 'C:/Users/DELL/Documents/DataFi/Data Review Meeting/Lab report Lookup to Radet/Filtered Lab report'
os.makedirs(output_base_dir, exist_ok=True)

# Define filter date
start_date = datetime(2024, 6, 20)
end_date =datetime(2025, 7, 31)

# Combine all CSV files into one DataFrame, specifying 'latin1' encoding
all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('.csv', '.xlsx'))]

# Combine all files into one DataFrame
combined_data = pd.DataFrame()

for file in all_files:
    try:
        if file.endswith('.csv'):
            data = pd.read_csv(file, encoding='latin1', on_bad_lines='skip')
        elif file.endswith('.xlsx'):
            data = pd.read_excel(file, engine='openpyxl')
        else:
            continue
        
        # Add Filename column
        data['Filename'] = os.path.basename(file)
        combined_data = pd.concat([combined_data, data], ignore_index=True)
    except Exception as e:
        print(f"Error processing file {file}: {e}")

if combined_data.empty:
    print("No valid files found or data could not be combined.")
    exit()

# Extract project name from the 'Filename' column
combined_data['ProjectName'] = combined_data['Filename'].str.split('_').str[0]
#combined_data['ProjectName'] = combined_data['IP']

# Converts date columns to datetime format
for col in ['Date Sample Collected (yyyy-mm-dd)']:
    
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Converts Last CD4 Count and Current Viral Load to integers
combined_data['Result'] = pd.to_numeric(combined_data['Result'], errors='coerce')

 


# Iterate through each unique ip
for project_name in combined_data['ProjectName'].unique():
    project_data = combined_data[combined_data['ProjectName'] == project_name].copy()

    # Create directory for the current ip
    project_dir = os.path.join(output_base_dir, project_name)
    os.makedirs(project_dir, exist_ok=True)

    
    all_line_lists_data = []

    # Condition for filtering
    Filter = project_data[(project_data['Test'] == 'Viral Load') &
                               ((project_data['Date Sample Collected (yyyy-mm-dd)'] >= start_date) & (project_data['Date Sample Collected (yyyy-mm-dd)'] <= end_date))]
    
    all_line_lists_data.append(Filter)

    

    # Combine all line list data for the current ip into one DataFrame
    all_line_lists_df = pd.concat(all_line_lists_data, ignore_index=True)

    # Define the output file path for the ip
    output_file_path = os.path.join(project_dir, f"{project_name}_Filtered.xlsx")

    # Save the transposed and the combined line list to one Excel file with two sheets
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        
        all_line_lists_df.to_excel(writer, sheet_name='Filtered Line List', index=False)

    print(f"Filtered for ip '{project_name}' saved to: {output_file_path}") 

