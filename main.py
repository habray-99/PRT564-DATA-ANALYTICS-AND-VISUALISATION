import pandas as pd
import numpy as np
import os
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score, mean_squared_error, mean_absolute_error
import warnings
warnings.filterwarnings('ignore')

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

        for sheet in sheets_to_process:
            df = pd.read_excel(file_path, sheet_name = sheet)
            print(f"  Sheet: {sheets_to_process}, Shape: {df.shape}")
            df = df.iloc[9:].reset_index(drop=True)

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


print("Preparing data for regression analysis...")

# Find columns with natural increase and migration data
natural_increase_cols = [col for col in final_df.columns if 'natural_increase' in col]
net_migration_cols = [col for col in final_df.columns if 'net_overseas_migration' in col]
net_interstate_cols = [col for col in final_df.columns if 'net_interstate_migration' in col]

print(f"Found {len(natural_increase_cols)} natural increase columns")
print(f"Found {len(net_migration_cols)} net overseas migration columns")
print(f"Found {len(net_interstate_cols)} net interstate migration columns")

# Create aggregated columns (Australia-wide)
model_data = pd.DataFrame({
    'date': final_df['date'],
    'year': final_df['year'],
    'quarter': final_df['quarter']
})

# Calculate Australia-wide averages
model_data['natural_increase_au'] = final_df[natural_increase_cols].mean(axis=1)
model_data['net_overseas_migration_au'] = final_df[net_migration_cols].mean(axis=1)
model_data['net_interstate_migration_au'] = final_df[net_interstate_cols].mean(axis=1)

# Remove rows with missing values
model_data = model_data.dropna()

print(f"\nModel data shape: {model_data.shape}")
print(f"\nFirst few rows:")
print(model_data.head(10))

if len(model_data) > 10 and not model_data[['natural_increase_au', 'net_overseas_migration_au']].isnull().all().all():
    # Prepare features and target
    X = model_data[['net_overseas_migration_au', 'net_interstate_migration_au']]
    y = model_data['natural_increase_au']
    
    # Train/test split
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    
    # Train model
    model = LinearRegression()
    model.fit(X_train, y_train)
    
    # Make predictions
    y_pred = model.predict(X_test)
    
    # Calculate metrics
    r2 = r2_score(y_test, y_pred)
    rmse = np.sqrt(mean_squared_error(y_test, y_pred))
    mae = mean_absolute_error(y_test, y_pred)
    
    print("="*60)
    print("ABS DATA REGRESSION ANALYSIS")
    print("="*60)
    print(f"\nModel Performance:")
    print(f"  R2 Score: {r2:.4f}")
    print(f"  RMSE: {rmse:.2f}")
    print(f"  MAE: {mae:.2f}")
    
    print(f"\nModel Coefficients:")
    for col, coef in zip(X.columns, model.coef_):
        print(f"  {col}: {coef:.4f}")
    print(f"  Intercept: {model.intercept_:.4f}")
    
else:
    print("Insufficient data for regression analysis.")
