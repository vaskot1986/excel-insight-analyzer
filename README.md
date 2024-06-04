# Excel Analyzer Tool

## Description
This project is an Excel Analyzer tool that processes multiple Excel files to extract and compile information about their columns. It includes functionalities to analyze sheet structures by a given marker, filter data, and gather column information.

## Files
- `main.py`: The main script to execute the analysis and processing of Excel files.
- `excel_analyzer.py`: Contains functions to gather information about columns in Excel sheets.
- `excel_processor.py`: Processes directories of Excel files, filtering and preparing them for analysis.

## Usage
1. **Analyzing and Filtering Excel Files**
   ```bash
   python main.py --analyze <path_to_excel_file> --output <output_excel_file>
   ```
   This command analyzes the specified Excel file and saves the column information to the output file.

2. **Processing Excel Files from a Directory**
   ```bash
   python main.py --path <directory_path> --output <output_excel_file>
   ```
   This command processes all Excel files in the specified directory and saves the aggregated column information to the output file.

## Requirements
- pandas
- openpyxl
- argparse
- os

## Installation
Ensure you have Python installed. You can install the required libraries using pip:
```bash
pip install pandas openpyxl
```

## Functions
### `main.py`
- **main**: The entry point of the script, parses arguments, and calls appropriate functions based on the arguments.
- **analyze_and_save**: Analyzes an Excel file and saves the column information.
- **process_directory_and_save**: Processes all Excel files in a directory and saves the column information.

### `excel_analyzer.py`
- **gather_column_info**: Gathers information about columns including data type, examples, and if the column is always empty.
- **save_column_info_to_excel**: Saves the gathered column information to an Excel file.

### `excel_processor.py**
- **load_and_filter_excel**: Loads and filters Excel files based on certain criteria.
- **analyze_sheet_structure_by_marker**: Analyzes the structure of an Excel sheet to identify headers based on a marker.

## Example
To analyze a specific Excel file and save the column information:
```bash
python main.py --analyze /path/to/excel/file.xlsx --output column_info_output.xlsx
```

To process all Excel files in a directory and save the aggregated column information:
```bash
python main.py --path /path/to/excel/files --output aggregated_column_info.xlsx
```

## License
This project is licensed under the MIT License.

