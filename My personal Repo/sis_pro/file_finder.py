import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def upload_file(file_num):
    filetypes = [("All files", "*.*")]
    filename = filedialog.askopenfilename(filetypes=filetypes)
    
    if filename:
        if filename.endswith('.xlsx'):
            file_type = 'xlsx'
        elif filename.endswith('.csv'):
            file_type = 'csv'
        else:
            messagebox.showerror("Error", "Unsupported file type. Please select a valid Excel (XLSX) or CSV file.")
            return

        if file_num == 1:
            entry_file1.delete(0, tk.END)
            entry_file1.insert(0, filename)
        else:
            entry_file2.delete(0, tk.END)
            entry_file2.insert(0, filename)

def compare_files():
    file1 = entry_file1.get()
    file2 = entry_file2.get()

    if not (file1 and file2):
        messagebox.showerror("Error", "Please select both files.")
        return

    try:
        # Determine file types
        file1_ext = os.path.splitext(file1)[1]
        file2_ext = os.path.splitext(file2)[1]

        # Read the two files based on their types
        if file1_ext.lower() == '.xlsx':
            excel1 = pd.read_excel(file1)
        elif file1_ext.lower() == '.csv':
            excel1 = pd.read_csv(file1)
        else:
            raise ValueError("Unsupported file format for File 1")

        if file2_ext.lower() == '.xlsx':
            excel2 = pd.read_excel(file2)
        elif file2_ext.lower() == '.csv':
            excel2 = pd.read_csv(file2)
        else:
            raise ValueError("Unsupported file format for File 2")

        # Get the number of columns in each dataframe
        num_cols = min(len(excel1.columns), len(excel2.columns))

        # Create a new DataFrame to store the result
        result = pd.DataFrame()

        # Iterate over pairs of columns and add them to the result DataFrame
        for i in range(0, num_cols, 2):
            col1_file1 = excel1.iloc[:, i]
            col1_file2 = excel2.iloc[:, i]

            # Add columns to the result dataframe
            result[f'Col_{i+1}_File1'] = col1_file1
            result[f'Col_{i+1}_File2'] = col1_file2
            result[f'Comparison_{i+1}'] = col1_file1.eq(col1_file2).map({True: 'Same', False: 'Different'})

            col2_file1 = excel1.iloc[:, i+1]
            col2_file2 = excel2.iloc[:, i+1]

            result[f'Col_{i+2}_File1'] = col2_file1
            result[f'Col_{i+2}_File2'] = col2_file2
            result[f'Comparison_{i+2}'] = col2_file1.eq(col2_file2).map({True: 'Same', False: 'Different'})

        # Determine the output file format based on the first input file
        if file1_ext.lower() == '.xlsx':
            output_file_ext = '.xlsx'
        elif file1_ext.lower() == '.csv':
            output_file_ext = '.csv'

        # Get the directory of the Python script
        script_directory = os.path.dirname(os.path.abspath(__file__))

        # Construct the path to save the result file
        result_file_path = os.path.join(script_directory, "result_with_comparison" + output_file_ext)

        # Write the result to a new file with the appropriate format
        if output_file_ext == '.xlsx':
            result.to_excel(result_file_path, index=False)
        elif output_file_ext == '.csv':
            result.to_csv(result_file_path, index=False)

        messagebox.showinfo("Success", f"Comparison completed. Results saved as '{result_file_path}'.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main Tkinter window
root = tk.Tk()
root.title("File Comparison")

# Create and place widgets
label_file1 = tk.Label(root, text="Select File 1:")
label_file1.grid(row=0, column=0, padx=5, pady=5)
entry_file1 = tk.Entry(root, width=50)
entry_file1.grid(row=0, column=1, padx=5, pady=5)
button_browse1 = tk.Button(root, text="Browse", command=lambda: upload_file(1))
button_browse1.grid(row=0, column=2, padx=5, pady=5)

label_file2 = tk.Label(root, text="Select File 2:")
label_file2.grid(row=1, column=0, padx=5, pady=5)
entry_file2 = tk.Entry(root, width=50)
entry_file2.grid(row=1, column=1, padx=5, pady=5)
button_browse2 = tk.Button(root, text="Browse", command=lambda: upload_file(2))
button_browse2.grid(row=1, column=2, padx=5, pady=5)

button_compare = tk.Button(root, text="Compare Files", command=compare_files)
button_compare.grid(row=2, column=1, padx=5, pady=5)

# Run the Tkinter event loop
root.mainloop()
