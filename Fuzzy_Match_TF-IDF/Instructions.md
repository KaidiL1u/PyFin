# Instructions for Using the Code
## User Manual: TF-IDF Text Matching Script

### Overview:
This script allows you to find the most similar text between two columns from an Excel file using the **TF-IDF** (Term Frequency-Inverse Document Frequency) method. It is designed to handle potentially large datasets and identify the closest matches based on the content of the text.

### Prerequisites:
- Python 3.x installed on your system.
- Required Python packages: `pandas`, `scikit-learn`, `tqdm`, `tkinter`, `chardet`, `concurrent.futures`, `sparse_dot_topn`.
  
  You can install these using pip:
  ```bash
  pip install pandas scikit-learn tqdm chardet sparse-dot-topn
