import tkinter as tk
from tkinter import filedialog, Toplevel, ttk
import pandas as pd
import os

def compare_excel_files(file1_path, file2_path, data_type, sort_column):
  """Compares two Excel files based on email IDs and extracts unique entries.

  Args:
    file1_path: Path to the first Excel file.
    file2_path: Path to the second Excel file.
    data_type: Whether to extract 'unique' or 'common' entries.
    sort_column: Column to sort the output by.
  """

  # Read the Excel files into Pandas DataFrames
  df1 = pd.read_excel(file1_path)
  df2 = pd.read_excel(file2_path)

  # Merge the DataFrames based on email IDs
  merged_df = pd.merge(df1, df2, on='email', how='left', indicator=True)

  # Filter based on data_type
  if data_type == 'unique':
    unique_df = merged_df[merged_df['_merge'] == 'left_only']
  elif data_type == 'common':
    unique_df = merged_df[merged_df['_merge'] == 'both']

  # Drop the '_merge' column and sort the output
  unique_df.drop('_merge', axis=1, inplace=True)
  unique_df.sort_values(by=sort_column, inplace=True)

  # Generate a unique output file name based on the input files
  output_file_path = f"output_{os.path.basename(file1_path)}_vs_{os.path.basename(file2_path)}.xlsx"

  # Write the unique entries to a new Excel file
  unique_df.to_excel(output_file_path, index=False)

  return output_file_path

def browse_file(entry):
  """Opens a file dialog for the user to select a file and updates the entry field."""
  file_path = filedialog.askopenfilename()
  entry.delete(0, tk.END)
  entry.insert(0, file_path)

def compare_files():
  """Compares the selected files and writes the unique entries to a new file."""
  file1_path = file1_entry.get()
  file2_path = file2_entry.get()

  if not file1_path or not file2_path:
    error_label.config(text="Please enter both file paths")
    return

  # Open a dialog box to get user input
  dialog = Toplevel(window)
  dialog.title("Comparison Options")

  data_type_label = tk.Label(dialog, text="Data Type:")
  data_type_var = tk.StringVar(value="unique")
  data_type_combo = ttk.Combobox(dialog, textvariable=data_type_var, values=["unique", "common"])
  data_type_combo.pack()

  sort_column_label = tk.Label(dialog, text="Sort Column:")
  sort_column_var = tk.StringVar()
  sort_column_combo = ttk.Combobox(dialog, textvariable=sort_column_var)
  sort_column_combo.pack()

  # Get common columns from both DataFrames
  df1 = pd.read_excel(file1_path)
  df2 = pd.read_excel(file2_path)
  common_columns = list(set(df1.columns) & set(df2.columns))
  sort_column_combo['values'] = common_columns

  ok_button = tk.Button(dialog, text="OK", command=lambda: dialog.destroy())
  ok_button.pack()

  dialog.wait_window()

  # Compare files with user-provided options
  output_file_path = compare_excel_files(file1_path, file2_path, data_type=data_type_var.get(), sort_column=sort_column_var.get())
  error_label.config(text=f"Comparison completed successfully! Output file: {output_file_path}")

# Create the main window
window = tk.Tk()
window.title("Excel File Comparison")

# Create labels and entry fields
file1_label = tk.Label(window, text="File 1:")
file1_entry = tk.Entry(window, width=50)
file1_button = tk.Button(window, text="Browse", command=lambda: browse_file(file1_entry))

file2_label = tk.Label(window, text="File 2:")
file2_entry = tk.Entry(window, width=50)
file2_button = tk.Button(window, text="Browse", command=lambda: browse_file(file2_entry))

compare_button = tk.Button(window, text="Compare", command=compare_files)
error_label = tk.Label(window, text="")

# Arrange widgets
file1_label.grid(row=0, column=0)
file1_entry.grid(row=0, column=1)
file1_button.grid(row=0, column=2)
file2_label.grid(row=1, column=0)
file2_entry.grid(row=1, column=1)
file2_button.grid(row=1, column=2)
compare_button.grid(row=2, column=1)
error_label.grid(row=3, column=1)

# Start the GUI
window.mainloop()
