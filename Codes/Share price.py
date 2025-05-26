import pandas as pd
import glob
import os

# Folder path containing CSV files (modify this as needed)
folder_path = "C:/Users/User/" # folder details need to add

# Get all CSV files in the folder
csv_files = glob.glob(os.path.join(folder_path, "*.csv"))

merged_data = []

for file in csv_files:
    df = pd.read_csv(file)  # Read CSV file
    df["Company Name"] = os.path.basename(file).replace("Stock Price History.csv", "")  # Extract company name
    merged_data.append(df)

# Combine all CSVs into a single DataFrame
final_df = pd.concat(merged_data, ignore_index=True)

# # Save the merged data to a new CSV file
# csv_output = "Stock values.csv"
# final_df.to_csv(csv_output, index=False)

# Convert CSV to Excel
excel_output = "Stock values.xlsx"
final_df.to_excel(excel_output, index=False, engine="openpyxl")

print(f"✅ Merging completed. The output files are  '{excel_output}'.")
