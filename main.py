import argparse
import pandas as pd
import os
from excel_analyzer import load_and_filter_excel, analyze_sheet_structure_by_marker, remove_trailing_none, gather_column_info, save_column_info_to_excel
from excel_processor import process_excel_files, save_to_excel

import argparse
import os
import pandas as pd

def main():
    parser = argparse.ArgumentParser(description="Excel Analyzer")
    parser.add_argument('--path', type=str, help='The directory to scan for Excel files')
    parser.add_argument("--output", type=str, help="Output Excel file to save the analysis")
    parser.add_argument("--analyze", type=str, help="File to load and filter")
    args = parser.parse_args()
    
    if args.path:
        directory = args.path
        data = process_excel_files(directory)
        save_to_excel(data, args.output)
        print(f"Done processing Excel files. Output saved to {args.output}")

    if args.analyze:
        filtered_df = load_and_filter_excel(args.analyze)
        total_documents = len(filtered_df)
        
        analysis_result = []
        column_info_list = []
        
        for index, row in filtered_df.iterrows():
            file_path = row['Filename']
            sheet_name = row['Sheet Name']
            marker = row['First Col']
            document_number = index + 1
            file_name = os.path.basename(file_path)  # Extract the filename
            
            try:
                headers, header_row_idx = analyze_sheet_structure_by_marker(file_path, sheet_name, marker)
                cleaned_headers = remove_trailing_none(headers)
                print(f"{document_number}/{total_documents} - Header for {file_path} - {sheet_name}: {cleaned_headers}")
                
                # Load the sheet into a DataFrame using the header_row_idx
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row_idx)
                
                # Gather column information
                col_info = gather_column_info(df, cleaned_headers)
                for column, info in col_info.items():
                    example1 = info['Examples'][0] if len(info['Examples']) > 0 else None
                    example2 = info['Examples'][1] if len(info['Examples']) > 1 else None
                    
                    print(f"Column: {column}")
                    print(f"  Data Type: {info['Data Type']}")
                    print(f"  Example1: {example1}")
                    print(f"  Example2: {example2}")
                    print(f"  Always Empty: {info['Always Empty']}")
                    
                    column_info_list.append({
                        "Column": column,
                        "Data Type(not-reliable)": info["Data Type"],
                        "Data Type(manual)": '',
                        "Rename To(manual)": '',
                        "Observations(manual)": '',
                        "Example1": example1,
                        "Example2": example2,
                        "Always Empty": info["Always Empty"],
                        "File": file_name,
                        "Sheet": sheet_name,
                    })
            except Exception as e:
                print(f"Error processing {file_path} - {sheet_name}: {e}")
        
        # Save column information to an Excel file
        column_info_file = args.output if args.output else 'column_info_output.xlsx'
        save_column_info_to_excel(column_info_list, column_info_file)
        print(f"Column information saved to {column_info_file}")

if __name__ == "__main__":
    main()
