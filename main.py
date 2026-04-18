import pandas as pd
import os

root_folder = os.getcwd()
dataset_folder = root_folder + "/dataset"
dataframes = []

for filename in os.listdir(dataset_folder):
    file_path = dataset_folder + "/" + filename
    print(file_path)
    try:
        print(f"Opening file {filename}")
        ind_file = pd.ExcelFile(file_path)
        print("------------------------------------------------------")
        sheets_to_process = [sheet for sheet in ind_file.sheet_names if sheet.startswith("Data")]
        print(sheets_to_process)

        for sheet in sheets_to_process:
            df = pd.read_excel(file_path, sheet_name = sheet, skiprows=9)
            # print(df)

            # First column contains the dates
            df_renamed = pd.DataFrame()
            df_renamed['date'] = df.iloc[:, 0]

            # Add metadata columns
            df_renamed['source_file'] = filename
            df_renamed['sheet_name'] = sheet

            
            # Parse date column
            df_renamed['date'] = pd.to_datetime(df_renamed['date'], errors='coerce')
            
            # Extract year from date
            df_renamed['year'] = df_renamed['date'].dt.year
            df_renamed['quarter'] = df_renamed['date'].dt.quarter
            
            # Extract all the measurement data
            for i, col in enumerate(df.columns[1:]):
                # Clean column name - extract measure and state
                col_clean = str(col).strip()
                
                # Parse the column header format: "Measure ; State ;"
                parts = [p.strip().strip(';') for p in col_clean.split(';') if p.strip()]
                
                if len(parts) >= 2:
                    measure = parts[0]
                    state = parts[1]
                else:
                    measure = col_clean
                    state = "Unknown"
                
                # Create a new column with the measurement
                col_name = f"{measure}_{state}".lower().replace(" ", "_")
                try:
                    df_renamed[col_name] = pd.to_numeric(df.iloc[:, i+1], errors='coerce')
                except:
                    pass
            
            dataframes.append(df_renamed)

            print(f"Finished retreiving data from {sheet}")
        
        print("------------------------------------------------------")
        print(f"closing file {filename}")
    except Exception as e:
        print(e)
        print('Above mentioned error has occured')

print(f"\nLoaded {len(dataframes)} sheets successfully!")


print("\nCombining data from all sheets...")

# Combine all dataframes (vertically stacking)
final_df = pd.concat(dataframes, ignore_index=True)

# Remove rows with missing dates
final_df = final_df.dropna(subset=['date', 'year'])

print(f"Combined data shape: {final_df.shape}")
print(f"Date range: {final_df['date'].min()} to {final_df['date'].max()}")

print(f"\nData loaded from:")
for filename in os.listdir(dataset_folder):
    file_data = final_df[final_df['source_file'] == filename]
    sheets = file_data['sheet_name'].unique() if 'sheet_name' in file_data.columns else ['Data1']
    print(f"  {filename}: {len(file_data)} rows from sheets {list(sheets)}")

print(f"\nTotal columns available: {len(final_df.columns)}")
print(f"First 5 rows:")
print(final_df.head())
