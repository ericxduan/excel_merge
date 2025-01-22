# %%
import os
import pandas as pd

folder_path = r'C:\Users\eduan1\OneDrive - KPMG\Desktop\Eric Duan\python\2025\merge_excel\input'
export_folder_path = r'C:\Users\eduan1\OneDrive - KPMG\Desktop\Eric Duan\python\2025\merge_excel\output'

def combine_excel_files(folder_path):
    # Get a list of all Excel files in the specified folder
    all_files = [ f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls') ]
    combined_df = pd.DataFrame()

    # Loop through each file and combine them
    for file in all_files:
        file_path = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(file_path)
            df[ 'FileName' ] = file

            # Adding LEFT 5 of the file name to first column (for LFS it would be the entity code)

            df[ 'Entity Code' ] = file[ :5 ]
            combined_df = pd.concat([ combined_df, df ], ignore_index=True)
                        
        except Exception as e:
            print(f"Error reading {file}: {e}")
    
    # Removing rows where the 8th Column has no value （if needed, change Column # and remove #hash for below two rows)
    # if not combined_df.empty and len(combined_df.columns) >= 8:
    #    combined_df = combined_df.dropna(subset=[ combined_df.columns[ 7 ] ])   

    return combined_df

combined_df = combine_excel_files(folder_path)

# Export the combined dataframe to the same folder
output_file = os.path.join(export_folder_path, 'combined_excel_file.xlsx')
combined_df.to_excel(output_file, index=False)

print(f"Combined Excel file saved to {output_file}")


