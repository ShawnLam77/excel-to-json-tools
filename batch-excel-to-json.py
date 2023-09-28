import pandas as pd  
import os  
  
# Specify the path to the folder where the Excel file is to be converted
folder_path = 'python-excel-to-json'  
  
# Get all Excel files in a folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]  
  
# Iterate through each Excel file for conversion 
for file_name in excel_files:  
    # Reading Excel files
    file_path = os.path.join(folder_path, file_name)  
    df = pd.read_excel(file_path)  
  
    # Convert DataFrame to JSON
    json_data = df.to_json(orient='records', indent=4)  
  
    # Create the corresponding JSON file path and filename
    json_file_path = os.path.join(folder_path, 'json_files')  
    os.makedirs(json_file_path, exist_ok=True)  # Creating a folder to save JSON files 
    json_file_name = os.path.splitext(file_name)[0] + '.json'  
    json_file_path = os.path.join(json_file_path, json_file_name)  
  
    # Save as JSON file
    with open(json_file_path, 'w', encoding='utf-8') as json_file:  
        json_file.write(json_data)  
  
    print(f'Successfully converted {file_name} to {json_file_name}.')  
  
print('Batch conversion completed.')  
