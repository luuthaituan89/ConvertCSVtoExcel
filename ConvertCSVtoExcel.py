import pandas as pd
import csv
import os
from tqdm import tqdm
from openpyxl import load_workbook


file_path = input("Please enter the path to the CSV file: ")


if not os.path.exists(file_path):
    print("Error: The file was not found. Please check the path and try again.")
elif not file_path.lower().endswith('.csv'):
    print("Error: The file is not in CSV format.")
else:
    try:
        with open(file_path, mode='r', encoding='utf-8') as file:
            dialect = csv.Sniffer().sniff(file.read(1024))
            file.seek(0)

            print("Converting CSV to Excel...")
            df = pd.read_csv(file, delimiter=dialect.delimiter)

            output_dir = input("Please enter the directory to save the Excel file: ")

            while not os.path.exists(output_dir) or not os.path.isdir(output_dir):
                print("Error: The directory does not exist. Please enter a valid directory.")
                output_dir = input("Please enter the directory to save the Excel file: ")

            output_file_name = input("Please enter the name for the Excel file (without extension): ")
            output_file = os.path.join(output_dir, f"{output_file_name}.xlsx")

            print("Saving data to Excel...")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for i in tqdm(range(len(df)), desc="Writing rows", unit="row"):
                    df.iloc[[i]].to_excel(writer, index=False, header=(i==0), startrow=i, sheet_name='Sheet1')

            wb = load_workbook(output_file)
            ws = wb.active

            print("Adjusting column widths...")
            for col in tqdm(ws.columns, desc="Processing columns", unit="col"):
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[col_letter].width = adjusted_width

            wb.save(output_file)
            print("Conversion complete. File saved as", output_file)
    except Exception as e:
        print("An error occurred:", e)
