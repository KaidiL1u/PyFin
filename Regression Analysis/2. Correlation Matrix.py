import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# -------------------------------------------
# User Manual: Correlation Matrix Generator
# -------------------------------------------
#
# Overview:
# This script calculates the correlation matrix for variables within an Excel sheet and saves it as a new Excel file.
# It is designed to test and analyze relationships between variables in a given dataset.
#
# How to Use:
# 1. Run the script.
# 2. A file dialog will open — select the Excel file you want to analyze (.xlsx only).
# 3. The script will read data from the "Sheet1" sheet.
# 4. The first column is excluded from the correlation analysis (assumed to be non-numeric or an identifier).
# 5. Another dialog will prompt you to save the output Excel file.
# 6. A confirmation message will appear after saving.
#
# Important Notes:
# - Sheet Name: Data must be in a sheet named "Sheet1". Modify 'sheet_name' if needed.
# - Supported Files: Only `.xlsx` format is supported.
# - Column Exclusion: The first column is excluded by default for correlation calculation.
#
# Error Handling:
# - Ensure files are not open or locked by another application during processing.
# - Check for correct file selection and save paths.
# -------------------------------------------


def select_file():
    """Open a file dialog to select an Excel file."""
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return file_path


def save_file():
    """Open a file dialog to choose where to save the output Excel file."""
    save_path = filedialog.asksaveasfilename(
        title="Save Correlation Matrix As",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return save_path


def calculate_correlation(file_path, save_path):
    """Load the Excel file, calculate the correlation matrix, and save the result."""
    try:
        # Load data
        df = pd.read_excel(file_path, sheet_name="Sheet1")
        
        # Exclude the first column
        x_columns = df.columns[1:]
        if x_columns.empty:
            raise ValueError("The sheet does not contain enough columns for correlation analysis.")
        
        # Calculate correlation matrix
        corr_matrix = df[x_columns].corr()

        # Save to Excel
        corr_matrix.to_excel(save_path, index=True)

        print("Correlation matrix has been successfully saved.")
        messagebox.showinfo("Success", "Correlation matrix has been saved successfully.")

    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")


def main():
    """Main function to execute the script."""
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window

    # Select the input file
    file_path = select_file()
    if not file_path:
        print("No file selected. Exiting.")
        return

    # Select the output save location
    save_path = save_file()
    if not save_path:
        print("No save path selected. Exiting.")
        return

    # Calculate and save the correlation matrix
    calculate_correlation(file_path, save_path)


if __name__ == "__main__":
    main()
