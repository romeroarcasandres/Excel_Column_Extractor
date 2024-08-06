import os
import pandas as pd
from tkinter import Tk, filedialog, simpledialog, messagebox
import string

def excel_col_to_num(col):
    """Convert Excel column letters to a 0-based column index"""
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num - 1

def main():
    # Hide the root Tkinter window
    root = Tk()
    root.withdraw()

    # Step 1: Prompt the user to select a directory where the .xlsx file is located
    directory = filedialog.askdirectory(title="Select Directory Containing .xlsx File")
    if not directory:
        messagebox.showerror("Error", "No directory selected. Exiting.")
        return

    # Get list of all Excel files in the directory
    excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    if not excel_files:
        messagebox.showerror("Error", "No Excel files found in the selected directory. Exiting.")
        return

    # For simplicity, we assume there's only one Excel file in the directory
    excel_file = os.path.join(directory, excel_files[0])
    base_name = os.path.splitext(os.path.basename(excel_file))[0]

    # Step 2: Prompt the user to indicate the columns he wants to export
    columns_input = simpledialog.askstring("Columns Input", 
                                           "Enter the columns to export, separated by commas (A,B,C...):")
    if not columns_input:
        messagebox.showerror("Error", "No columns specified. Exiting.")
        return

    columns_letters = [col.strip() for col in columns_input.split(',')]
    columns_indices = [excel_col_to_num(col) for col in columns_letters]

    # Step 3: Prompt the user to select the output format
    output_format = simpledialog.askstring("Output Format", 
                                           "Enter the output format: 'txt' for tab-separated text file or 'xlsx' for Excel file:",
                                           initialvalue="xlsx").lower()
    if output_format not in ['txt', 'xlsx']:
        messagebox.showerror("Error", "Invalid format specified. Exiting.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)

        # Convert column indices to column names
        columns_names = df.columns[columns_indices]

        # Generate the output file with the specified columns
        output_df = df[columns_names]
        output_base = os.path.join(directory, base_name + "_columns")
        
        if output_format == 'txt':
            output_file = output_base + ".txt"
            output_df.to_csv(output_file, sep='\t', index=False)
        elif output_format == 'xlsx':
            output_file = output_base + ".xlsx"
            output_df.to_excel(output_file, index=False)
        else:
            messagebox.showerror("Error", "Invalid format specified. Exiting.")
            return

        messagebox.showinfo("Success", f"Output file generated: {output_file}")
        
    except IndexError:
        messagebox.showerror("Error", "The specified column is out of range. Please enter a column that exists in the Excel file.")
    except KeyError as e:
        messagebox.showerror("Error", f"One or more columns specified do not exist in the Excel file. {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    main()
