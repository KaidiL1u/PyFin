import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from tqdm import tqdm
from tkinter import filedialog, Tk, messagebox
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from sparse_dot_topn import awesome_cossim_topn

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


# Set seed for reproducibility
random_seed = 42

# Set lower bound of similarity to consider as a match
similarity_lower_bound = 0.1

# Set the number of characters in each n-gram
n_gram = 3

def clean_string(string):
    """Clean the string by removing any whitespace and non-alphanumeric characters."""
    if isinstance(string, str):
        return re.sub(r'\s+|[^a-zA-Z0-9]', '', string)
    else:
        return ''

def ngrams(string, n=n_gram):
    """Convert the string into a list of n-grams containing only alphanumeric characters."""
    return [''.join(ngram) for ngram in zip(*[string[i:] for i in range(n)])]

def process_match(row_index, matches, pbar):
    """
    Process the matches for a particular row by sorting them based on similarity score
    and returning the best match along with its similarity score.
    """
    try:
        if matches:
            matches.sort(key=lambda x: x[1], reverse=True)  # Sort matches in descending order of similarity score
            col_index, score = matches[0]  # Get the index and similarity score of the best match
            result = {
                short_label: df_short['original'].iloc[row_index],  # Get the original short text
                f'Best Match {long_label}': df_long['original'].iloc[col_index],  # Get the original long text of the best match
                'Similarity': score  # Get the similarity score of the best match
            }
        else:
            result = {
                short_label: df_short['original'].iloc[row_index],  # Get the original short text
                f'Best Match {long_label}': '1NoMatch',  # Indicate no match found
                'Similarity': None  # No similarity score when no match
            }
        pbar.update(1)  # Update the progress bar
        return result
    except Exception as e:
        pbar.update(1)
        return {
            short_label: df_short['original'].iloc[row_index],
            f'Best Match {long_label}': f'Error: {str(e)}',
            'Similarity': None
        }

if __name__ == "__main__":
    # Create top priority pop-up window
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # Select the file to read
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="Select the Excel file to read")

    if not file_path:
        messagebox.showerror("Error", "No file selected. Exiting.")
        exit()

    # Save the match results to an XLSX file
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                                    title="Choose the location and filename to save the output file")

    if not output_file_path:
        messagebox.showinfo("File Save Cancelled", "File save cancelled. Exiting.")
        exit()

    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path, sheet_name="Sheet1", usecols=[0, 1])

    # Extract the dynamic labels from the first row (A1 and B1)
    short_label = df.columns[0]
    long_label = df.columns[1]

    # Clean the data by removing any whitespace and non-alphanumeric characters, and store the original strings
    df_short = pd.DataFrame()
    df_short['cleaned'] = df.iloc[:, 0].apply(clean_string)
    df_short['original'] = df.iloc[:, 0]

    df_long = pd.DataFrame()
    df_long['cleaned'] = df.iloc[:, 1].apply(clean_string)
    df_long['original'] = df.iloc[:, 1]

    # Filter out blank rows in df_short
    df_short = df_short[df_short['cleaned'].str.strip() != ''].reset_index(drop=True)

    # Vectorize the texts using TF-IDF (Term Frequency-Inverse Document Frequency)
    vectorizer = TfidfVectorizer(min_df=1, analyzer=ngrams)  # Create a TF-IDF vectorizer with n-gram analyzer
    vectorizer.fit(df_long['cleaned'].astype(str))  # Fit the vectorizer only to the texts in the long column

    tf_idf_matrix_long = vectorizer.transform(df_long['cleaned'])  # Transform long texts into TF-IDF matrix
    tf_idf_matrix_short = vectorizer.transform(df_short['cleaned'])  # Transform short texts into TF-IDF matrix

    # Calculate the total number of iterations for matching (limited to the number of non-blank records in Column A)
    total_iterations = df_short.shape[0]

    # Use sparse_dot_topn library to efficiently find the top matches based on cosine similarity
    matches = awesome_cossim_topn(tf_idf_matrix_short, tf_idf_matrix_long.transpose(), ntop=1,
                                  use_threads=True, lower_bound=similarity_lower_bound)

    # Construct a dictionary to store matches for each short text
    match_dict = {}
    for row, col, score in zip(*matches.nonzero(), matches.data):
        match_dict.setdefault(row, []).append((col, score))

    # Display progress bar in advance
    with tqdm(total=total_iterations, desc="Extracting Matches", position=0, dynamic_ncols=True, unit_scale=False, unit=' rows') as pbar:
        # Use ThreadPoolExecutor to process matches in parallel
        with ThreadPoolExecutor(max_workers=None) as executor:
            futures = [executor.submit(process_match, row_index, matches, pbar) for row_index, matches in match_dict.items()]

            output_data = []
            for future in as_completed(futures):
                result = future.result()
                output_data.append(result)  # Collect processed matches

    # Handle rows in df_short that did not have any matches
    unmatched_rows = set(range(df_short.shape[0])) - set(match_dict.keys())
    for row_index in unmatched_rows:
        output_data.append({
            short_label: df_short['original'].iloc[row_index],  # Get the original short text
            f'Best Match {long_label}': '1NoMatch',  # No match found
            'Similarity': None  # No similarity score
        })

    # Create a DataFrame from the processed matches
    output_df = pd.DataFrame(output_data)

    # Save the match results to the specified XLSX file
    output_df.to_excel(output_file_path, index=False)  # Save DataFrame to XLSX file
    messagebox.showinfo("File Saved", f"File saved to: {output_file_path}")  # Inform the user that the file has been saved
