import os
import pandas as pd

def find_excel_files(directory):
    """Scan the directory for Excel files and return a list of their paths."""
    excel_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx'):
                excel_files.append(os.path.join(root, file))
    return excel_files

def read_excel_file(file_path):
    """Read an Excel file and return its content (simplified for demonstration)."""
    # This function should be expanded based on how you need to process each Excel file.
    df = pd.read_excel(file_path)
    return df

def list_excel_sheets(file_path):
    """List all sheet names in an Excel file."""
    excel_file = pd.ExcelFile(file_path)
    return excel_file.sheet_names

def process_excel_files(directory):
    """Main function to process all Excel files in a directory and create a summary Excel file."""
    excel_files = find_excel_files(directory)
    data = []

    for file_path in excel_files:
        sheet_names = list_excel_sheets(file_path)
        for sheet in sheet_names:
            data.append({'Filename': file_path, 'Sheet Name': sheet, 'First Col':'', 'Ignore':''})
    
    return data

def save_to_excel(data, output_path):
    """Save the collected data to an Excel file."""
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)