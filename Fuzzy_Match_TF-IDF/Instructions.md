# Instructions for Using the Code
# User Manual: TF-IDF Text Matching Script

# Overview:
# This script allows you to find the most similar text between two columns from an Excel file using
# the TF-IDF (Term Frequency-Inverse Document Frequency) method. It is designed to handle potentially
# large datasets and identify the closest matches based on the content of the text.

# Prerequisites:
# - Python 3.x installed on your system.
# - Required Python packages: `pandas`, `scikit-learn`, `tqdm`, `tkinter`, `chardet`, `concurrent.futures`, `sparse_dot_topn`.
#   You can install these using pip:
#   ```
#   pip install pandas scikit-learn tqdm chardet sparse_dot_topn
#   ```

# How to Use:
# 1. **Run the Script**:
#    - When you run the script, a window will pop up asking you to select an Excel file. This file should have
#      the text data you want to compare, with the data located in the first two columns (A and B) of "Sheet1".
#
# 2. **Save the Results**:
#    - After selecting the input file, another window will prompt you to choose a location and filename to save the results.
#      The output will be saved in an Excel file with the results of the text matching.
#
# 3. **Check the Results**:
#    - The output file will contain three columns:
#      - The original text from Column A.
#      - The most similar text found from Column B.
#      - The similarity score indicating how closely the texts match.
#    - If no match is found, "1NoMatch" will be displayed in the corresponding cell.

# Important Notes:
# - **Column Layout**: Ensure that your data is in the first two columns (A and B) of "Sheet1" in the Excel file.
# - **Handling Blank Cells**: The script will automatically ignore any blank cells in Column A during the matching process.
# - **Data Cleaning**: The script will remove any non-alphanumeric characters and whitespace before performing the matching.
# - **Matching Limitations**: The script will only attempt to find matches for non-blank cells in Column A against all rows in Column B.
# - **Performance**: The script uses multithreading to improve performance, but processing large files may still take some time.


# Please ensure that you install below all python package
pip install pandas scikit-learn tqdm sparse-dot-topn


