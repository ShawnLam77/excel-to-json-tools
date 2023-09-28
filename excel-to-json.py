import pandas as pd  
import os  
  
# Specify the path to the Excel file to be converted  
excel_file_path = 'python-excel-to-json/Sam.ple2.xlsx'  
  
# Read the Excel file  
df = pd.read_excel(excel_file_path)  
  
# Convert to JSON  
json_data = df.to_json(orient='records', indent=4)  
  
# Get the directory and filename of the Excel file  
file_dir, file_name = os.path.split(excel_file_path)  
file_name_without_ext = os.path.splitext(file_name)[0]  
  
# Create the JSON file path and filename  
json_file_path = os.path.join(file_dir, file_name_without_ext + '.json')  
  
# Save as JSON file  
with open(json_file_path, 'w', encoding='utf-8') as json_file:  
    json_file.write(json_data)  
  
print(f'Successfully converted {excel_file_path} to {json_file_path}.')  
