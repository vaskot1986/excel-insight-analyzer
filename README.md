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
   The output file contains the columns "first col" where the user need to write manually the first column header of each sheet. This way on the next step the script can identify where to start looking for the headers. Column "ignore" if set to "YES" will ignore the whole row. Use it to ignore complete sheets.

2. **Processing Excel Files from a Directory**
   ```bash
   python main.py --path <directory_path> --output <output_excel_file>
   ```
   This command processes all Excel files in the specified directory and saves the aggregated column information to the output file.
   There are 3 columns that can be manually filled and that are focused in helping the person that will use this document later.
   The colum DataType needs to be manually set, use the example columns to find out the data type. The colum rename can be used to define a rename for a specific column. The column observations can inlcude observations such as format, type of object, etc.

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
### Installation in Mac (pip & homebrew)
Setting up python virtual environment for python and pip lib management in mac using homebrew :

1. **Create a Virtual Environment**
   ```bash
   python3 -m venv ~/excel_parser/venv
   ```

2. **Activate the Virtual Environment**
   ```bash
   source ~/excel_parser/venv/bin/activate
   ```
3. **Upgrade pip Within the Virtual Environment**
   ```bash
   python -m pip install --upgrade pip
   ```
4. **Install Packages**
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

Copyright (c) 2024 Arnau Vazquez (github.com/vaskot1986)

This project is licensed under the MIT License. See the LICENSE file for details.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Note: This software is provided 'as is', and users are allowed to use, copy,
modify, merge, publish, distribute, sublicense, and/or sell copies of the
Software at their own risk.

