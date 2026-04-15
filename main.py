import pandas as pd
import os

root_folder = os.getcwd()
dataset_folder = root_folder + "/dataset"

for filename in os.listdir(dataset_folder):
    file_path = dataset_folder + "/" + filename
    print(file_path)
    print(f"Opening file {filename}")
    print("------------------------------------------------------")
    df = pd.read_excel(file_path, skiprows= 1)
    print(df)
    print("------------------------------------------------------")
    print(f"closinf file {filename}")
