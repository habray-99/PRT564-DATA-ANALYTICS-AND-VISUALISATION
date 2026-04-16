import pandas as pd
import os

root_folder = os.getcwd()
dataset_folder = root_folder + "/dataset"

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
            print(df)

            # First column contains the dates
            df_renamed = pd.DataFrame()
            df_renamed['date'] = df.iloc[:, 0]


            print(f"Finished retreiving data from {sheet}")
        
        print("------------------------------------------------------")
        print(f"closing file {filename}")
    except Exception as e:
        print(e)
        print('file cannot be found')
