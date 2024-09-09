# main.py

from core import convert_excel_to_json, query_data

def main():
    file_path = input("Enter the path to the Excel file: ")
    sheet_columns, data = convert_excel_to_json(file_path)
    if sheet_columns and data:
        query_data(sheet_columns, data)

if __name__ == "__main__":
    main()
