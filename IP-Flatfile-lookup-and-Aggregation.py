import pandas as pd
import os

# folder and file paths
ip_flatfile_folder = 'C:/Users/DELL/Documents/DataFi/IP_Flatfile_Aggregation & Analysis/IP_Flatfile2'
codelist_file = 'C:/Users/DELL/Documents/DataFi/IP_Flatfile_Aggregation & Analysis/codelist/codelist(fac & comm).xlsx'
org_unit_codelist_file = 'C:/Users/DELL/Documents/DataFi/IP_Flatfile_Aggregation & Analysis/codelist/org_unit_codelist.xlsx'
coc_codelist_file = 'C:/Users/DELL/Documents/DataFi/IP_Flatfile_Aggregation & Analysis/codelist/COC codelist(fac & comm)2.xlsx'
output_lookedup_folder = 'C:/Users/DELL/Documents/DataFi/IP_Flatfile_Aggregation & Analysis/looked_upQ2b'
output_aggregated_file = 'C:/Users/DELL/Documents/DataFi/IP_Flatfile_Aggregation & Analysis/Aggregated_IP_Flatfile_data_updatedALL2.xlsx' #add 2 when lookingup for COC


# Creates output folders if it doesn't exist
os.makedirs(output_lookedup_folder, exist_ok=True)

# Load
try:
    codelist_df = pd.read_excel(codelist_file)
    org_unit_codelist_df = pd.read_excel(org_unit_codelist_file)
    coc_codelist_df = pd.read_excel(coc_codelist_file)
except FileNotFoundError as e:
    print(f"Error: One or more codelist files not found: {e}")
    exit()

# Process each file in the IP_Flatfile folder
for filename in os.listdir(ip_flatfile_folder):
    if filename.endswith(('.csv', '.xlsx')):
        filepath = os.path.join(ip_flatfile_folder, filename)
        ip_name = filename.split('_')[0]

        try:
            if filename.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)

            # Add filename and IP_name columns
            df['filename'] = filename
            df['IP_name'] = ip_name

            # Lookup from codelist(fac & comm)
            df = pd.merge(df, codelist_df[['dataelement', 'shortname',  'Indicator']], #uncomment when lookingup for COC
                          on='dataelement', how='left')
            df.drop(columns=['dataelement'], inplace=True)

            #Lookup from org_unit_codelist
            df = pd.merge(df, org_unit_codelist_df[['orgUnit', 'orgunit_parent']],  #uncomment when lookingup for COC
                          on='orgUnit', how='left')
            #df.drop(columns=['orgunit_parent_internal_id'], inplace=True)

            # df = pd.merge(df, coc_codelist_df[['categoryoptioncombo', 'Indicator2']], #uncomment when lookingup for DE name
            #               on='categoryoptioncombo', how='left')



            

            # Save the looked-up data to the new folder
            output_filepath = os.path.join(output_lookedup_folder, filename)
            if filename.endswith('.csv'):
                df.to_csv(output_filepath, index=False)
            else:
                df.to_excel(output_filepath, index=False)

            print(f"Processed and looked up data for: {filename}")

        except Exception as e:
            print(f"Error processing file {filename}: {e}")

print("\nLookup process complete. Files with looked-up data are in:", output_lookedup_folder)

# --- Aggregation/Pivot Process ---
all_data = {}     #all_data = []
for filename in os.listdir(output_lookedup_folder):
    if filename.endswith(('.csv', '.xlsx')):
        filepath = os.path.join(output_lookedup_folder, filename)
        ip_name = filename.split('_')[0]

        try:
            if filename.endswith('.csv'):
                df_lookedup = pd.read_csv(filepath, encoding='latin1', on_bad_lines='skip')
            else:
                df_lookedup = pd.read_excel(filepath, engine='openpyxl')

            # Pivot the data by Indicator and value
            if 'Indicator' in df_lookedup.columns and 'value' in df_lookedup.columns:        #change to Indicator2 when lookingup for COC
                pivot_df = df_lookedup.pivot_table(index='Indicator', columns='IP_name', values='value', aggfunc='sum')  #change to Indicator2 when lookingup for COC
                all_data[ip_name] = pivot_df    #all_data.append(pivot_df)
            else:
                print(f"Warning: 'Indicator' or 'value' column not found in {filename} for pivoting.")

        except Exception as e:
            print(f"Error reading looked-up file {filename} for aggregation: {e}")


# Save all pivoted dataframes to different sheets in the same Excel workbook
if all_data:
    with pd.ExcelWriter(output_aggregated_file, engine='openpyxl') as writer:
        for ip_name, df_pivot in all_data.items():
            df_pivot.to_excel(writer, sheet_name=ip_name, index=True) 
    print("\nAggregation process complete. Aggregated data saved to different sheets in:", output_aggregated_file)
else:
    print("\nNo data was aggregated.")

# # Concatenate all pivoted dataframes
# if all_data:
#     aggregated_df = pd.concat(all_data, ignore_index=False)

#     # Save the aggregated data
#     aggregated_df.to_excel(f"{output_aggregated_file}.xlsx")
#     print("\nAggregation process complete. Aggregated data saved to:", f"{output_aggregated_file}.xlsx")
# else:
#     print("\nNo data was aggregated.")

