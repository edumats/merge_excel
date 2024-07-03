import os

import pandas as pd
import xlrd

excel_files = []

# Function to generate a unique filename
def generate_unique_filename(filename):
    base, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(new_filename):
        new_filename = f"{base}_{counter}{extension}"
        counter += 1
    return new_filename

# Extract the folder name
def extract_folder_name(directory):
    return os.path.basename(os.path.normpath(directory))

def main():
    directory = input('Excel files directory: ')

    for filename in os.listdir(directory):
        print('Filename', filename)
        if filename.endswith(".xls"):
            file_path = os.path.join(directory, filename)
            try:
                workbook = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
                df = pd.read_excel(workbook)
                excel_files.append(df)
            except Exception as e:
                print('Error reading file:', filename, e)

    if excel_files:
        consolidated_excel = pd.concat(excel_files, ignore_index=True)

        folder_name = extract_folder_name(directory)

        output_filename = generate_unique_filename(f'{folder_name}.xlsx')

        consolidated_excel.to_excel(output_filename, index=False, engine='openpyxl')
    else:
        print('No Excel files found in the directory')

if __name__ == '__main__':
    main()