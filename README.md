# Excel_Column_Extractor
This script extracts specified columns from an Excel (.xlsx) file and exports the data into either a tab-separated text file or a new Excel file.

## Overview:
Column_Extractor allows users to select specific columns from an Excel file by specifying column letters (e.g., A, B, Z, AB, CD). The user can export group of columns in different files using the semicolon (e.g., A,B,E; C,D,F) 
The script prompts the user to choose the desired columns and the output format, either as a .txt file or another .xlsx file.
It is particularly useful for quickly extracting subsets of data for further analysis or processing.

## Requirements
* Python 3
* pandas library (for data manipulation)
* tkinter library (for user interface)
* os library (for file path operations)
* string library (for handling string operations)

## Files
column_extractor.py

## Usage
1. Run the script.
2. A file dialog will prompt you to select a directory containing the Excel (.xlsx) file.
3. A dialog box will prompt you to select the desired columns (or group of columns) by specifying their letters (e.g., A,B,E; C,D,F).
4. Another dialog box will prompt you to select the desired output format:
   *  "txt" for a tab-separated text file.
   *  "xlsx" for a new Excel file.
5. The output file will be saved in the same directory as the input file, with an appended "_columns" and the column groups to the original file name.

## Important Notes
* It is recommended that the selected directory contains only one Excel file to avoid confusion.
* The script supports Excel files with multiple sheets but will only process the first sheet by default.
* The columns will be exported in the order indicated. Each group of columns will be exported in the order indicated and saved as separate files.
* The column specification should match the Excel sheet's structure; ensure the columns exist within the file.
* The output format choice must be either 'txt' or 'xlsx'; any other format will result in an error.

## License
This project is governed by the CC BY-NC 4.0 license. For comprehensive details, kindly refer to the LICENSE file included with this project.
