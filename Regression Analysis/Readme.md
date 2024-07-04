**Required Python Packages**

To ensure the code runs successfully, please install the following Python packages:
# Install pandas
pip install pandas
# Install statsmodels
pip install statsmodels
# Install numpy
pip install numpy
# Install tkintertable
pip install tkintertable
# Install openpyxl
pip install openpyxl


**# Instructions for Using the Code**
The Python code reads an .xlsx file and uses "Sheet1" as the source data.

Ensure the following structure for the data in "Sheet1":

Cell A1 should be named 'Year' and contain the related years for each variable. Column A is the column for years.
Column B should contain your Y variable for the regression analysis.
Columns C and beyond should contain the X variables.
Do not leave any blank columns or rows in between to avoid crashing.
Row 1 should contain the headers for each variable. Do not leave any cells blank. Carefully name each of your variables, including:

Cell A1: 'Year'
Cell B1: The name of the Y variable
Cell C1 and beyond: The names of the X variables
