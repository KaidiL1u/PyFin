import pandas as pd
import tkinter as tk
from tkinter import filedialog

# User Manual: Correlation Matrix Generator Script
#
# Overview:
# This script calculates the correlation matrix for variables within an Excel sheet and saves it as a new Excel file.
# It is designed to test and analyze the relationships between variables in the given dataset.
#
# How to Use:
# 
# 1. Run the Script:
#    - The script will open a file dialog to select an Excel file. Choose the file you want to analyze.
#
# 2. File Selection:
#    - A file dialog will appear for you to select an Excel file. Only `.xlsx` files are supported.
#    - After selecting a file, the script will load data from the sheet named "Sheet1".
#
# 3. Data Handling:
#    - The script assumes the first column in the sheet is not part of the correlation analysis and excludes it.
#      It calculates correlations for the remaining columns.
#
# 4. Save the Output:
#    - A save dialog will prompt you to select a location and name for the output file.
#    - The correlation matrix will be saved as an Excel file at the specified location.
#
# 5. Confirmation:
#    - After saving the file, a confirmation message will be printed to indicate that the correlation matrix has been successfully saved.
#
# Important Notes:
# 
# - Sheet Name: Ensure the sheet you want to analyze is named "Sheet1". If your data is in a different sheet, modify the script accordingly.
# - File Format: The script only supports `.xlsx` files. Ensure your input file is in the correct format.
# - Column Exclusion: The script excludes the first column of the data for correlation calculations, assuming it might be a non-numeric or identifier column.
#   Make sure this aligns with your data structure.
#
# Error Handling:
# 
# - If any issues occur during file selection or saving, check the file dialogs and paths provided. Make sure the Excel file is not open or locked by another application during processing.



# Initialize the main Tkinter window
root = tk.Tk()
root.withdraw()  # Hide the root window

# Function to bring file dialogs to the front
def make_window_top_priority(window):
    window.attributes('-topmost', True)  # Set the window to be on top
    window.attributes('-topmost', False)  # Then reset the topmost attribute

# Ask user to select the Excel file
file_dialog = tk.Toplevel()  # Create a new top-level window for the file dialog
make_window_top_priority(file_dialog)  # Bring the file dialog to the front
file_path = filedialog.askopenfilename(
    parent=file_dialog,  # Set the file dialog's parent to the top-level window
    title="Select Excel file",  # Title of the file dialog
    filetypes=[("Excel files", "*.xlsx")]  # Only allow Excel files to be selected
)
file_dialog.destroy()  # Destroy the file dialog window after selection

# Load data from the selected Excel file into a dataframe
df = pd.read_excel(file_path, sheet_name="Sheet1")  # Read the specified sheet into a dataframe

# Exclude the first column (assuming it's year) and set the rest as X variables
x_columns = df.columns[1:]  # Get all column names except the first one

# Calculate the correlation matrix for the X variables
corr_matrix = df[x_columns].corr()  # Compute the correlation matrix for the selected columns

# Ask user to choose location and name to save the Excel file
save_dialog = tk.Toplevel()  # Create a new top-level window for the save dialog
make_window_top_priority(save_dialog)  # Bring the save dialog to the front
save_path = filedialog.asksaveasfilename(
    parent=save_dialog,  # Set the save dialog's parent to the top-level window
    defaultextension=".xlsx",  # Set default file extension
    filetypes=[("Excel files", "*.xlsx")]  # Only allow Excel files to be saved
)
save_dialog.destroy()  # Destroy the save dialog window after selection

# Export the correlation matrix to an Excel file
corr_matrix.to_excel(save_path, index=True)  # Save the correlation matrix to the specified path

print("Correlation matrix for X variables has been saved as an Excel file.")  # Print confirmation message

