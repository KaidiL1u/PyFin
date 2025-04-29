import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import statsmodels.api as sm
import numpy as np
import sys
import warnings
from tkintertable import TableCanvas, TableModel
from openpyxl import Workbook

# Suppressing FutureWarnings regarding pandas deprecations
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=UserWarning)

# Global counter for the number of regressions run
regression_counter = 0


def read_xlsx_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        root.destroy()
        create_main_window(df)


def create_main_window(df):
    root = tk.Tk()
    root.title("Select Variables and Options")
    root.geometry("800x900")
    y_variable_name = df.columns[1]  # Dynamically fetching the Y variable name from column B
    year_frame, main_frame = create_selection_frames(root)
    year_vars = create_year_checkboxes(year_frame, df['Year'].unique())
    var_dict = create_variable_checkboxes(main_frame, df.columns[2:])
    force_intercept_var = tk.BooleanVar(value=False)
    tk.Checkbutton(main_frame, text="Force Intercept to Zero", variable=force_intercept_var).pack(side='top', padx=10,
                                                                                                  pady=(20, 0))
    create_action_buttons(main_frame, df, year_vars, var_dict, force_intercept_var, y_variable_name)
    root.mainloop()


def create_selection_frames(root):
    year_frame = tk.Frame(root)
    year_frame.pack(side='left', padx=10, pady=10, fill='y')
    tk.Label(year_frame, text="Select Years", font=('Helvetica', 14, 'bold')).pack(pady=(0, 10))
    separator = tk.Frame(root, width=2, bg='gray')
    separator.pack(side='left', fill='y', padx=10, pady=10)
    main_frame = tk.Frame(root)
    main_frame.pack(side='right', padx=10, pady=10, fill='both', expand=True)
    return year_frame, main_frame


def create_year_checkboxes(year_frame, years):
    year_vars = {year: tk.BooleanVar() for year in years}
    for year, var in year_vars.items():
        chk = tk.Checkbutton(year_frame, text=year, variable=var)
        chk.pack(anchor='w')
    # Add buttons for selecting and clearing all years
    tk.Button(year_frame, text="Select All Years", command=lambda: update_checkboxes(year_vars, True)).pack(side='top',
                                                                                                            padx=10,
                                                                                                            pady=5)
    tk.Button(year_frame, text="Clear All Years", command=lambda: update_checkboxes(year_vars, False)).pack(side='top',
                                                                                                            padx=10,
                                                                                                            pady=5)
    return year_vars


def create_variable_checkboxes(main_frame, variables):
    var_dict = {var: tk.BooleanVar() for var in variables}
    for var, variable in var_dict.items():
        checkbox = tk.Checkbutton(main_frame, text=var, variable=variable)
        checkbox.pack(anchor='w')
    return var_dict


def update_checkboxes(variable_dict, value):
    for var in variable_dict.values():
        var.set(value)


def create_action_buttons(main_frame, df, year_vars, var_dict, force_intercept_var, y_variable_name):
    tk.Button(main_frame, text="Run Regression Test",
              command=lambda: run_analysis(df, year_vars, var_dict, force_intercept_var.get(), y_variable_name)).pack(
        side='top', padx=10, pady=20)
    tk.Button(main_frame, text="Quit", command=sys.exit).pack(side='right', padx=10, pady=10)


def run_analysis(df, year_vars, var_dict, force_intercept, y_variable_name):
    global regression_counter
    regression_counter += 1
    selected_years = [year for year, var in year_vars.items() if var.get()]
    selected_x_vars = [var for var, val in var_dict.items() if val.get()]
    df_selected = df[df['Year'].isin(selected_years)]
    columns_to_keep = ['Year', y_variable_name] + selected_x_vars
    df_selected = df_selected[columns_to_keep]
    model = run_regression(df_selected, force_intercept, y_variable_name)
    if model:
        output_df = format_regression_output(model)
        anova_table = calculate_anova_table(model)
        show_results_window(output_df, selected_years, y_variable_name, regression_counter, model, anova_table)


def run_regression(df, force_intercept, y_variable_name):
    Y = df[y_variable_name].astype(float)
    X = df[df.columns.difference(['Year', y_variable_name])].astype(float)
    if not force_intercept:
        X = sm.add_constant(X)
    model = sm.OLS(Y, X).fit()
    return model


def format_regression_output(model):
    summary_df = pd.read_html(model.summary().as_html(), header=0, index_col=0)[1]
    return summary_df


def calculate_anova_table(model):
    sse = model.ssr  # Sum of squared residuals
    ssr = model.ess  # Explained sum of squares
    sst = ssr + sse  # Total sum of squares
    dfe = model.df_resid  # Degrees of freedom for error
    dfr = model.df_model  # Degrees of freedom for regression
    dft = dfr + dfe  # Total degrees of freedom

    mse = sse / dfe  # Mean squared error
    msr = ssr / dfr  # Mean squared regression

    f_stat = msr / mse  # F-statistic
    p_value = model.f_pvalue  # P-value for the F-statistic

    anova_table = pd.DataFrame({
        'df': [dfr, dfe, dft],
        'SS': [ssr, sse, sst],
        'MS': [msr, mse, np.nan],
        'F': [f_stat, np.nan, np.nan],
        'Significance F': [p_value, np.nan, np.nan]
    }, index=['Regression', 'Residual', 'Total'])

    return anova_table


def show_results_window(output_df, selected_years, y_variable_name, run_number, model, anova_table):
    results_window = tk.Toplevel()
    results_window.title(f"Regression Results - Run {run_number}: {y_variable_name}")
    results_window.geometry("1200x700")

    frame = tk.Frame(results_window)
    frame.pack(fill='both', expand=True)

    # Prepare data for the table
    summary_data = []

    # Add selected years at the top
    summary_data.append(['Selected Years', ', '.join(map(str, selected_years))])

    # Add regression statistics
    summary_data.append(['SUMMARY OUTPUT', ''])
    summary_data.append([''])
    summary_data.append(['Regression Statistics', ''])
    summary_data.append(['Multiple R', f"{model.rsquared ** 0.5:.4f}"])
    summary_data.append(['R Square', f"{model.rsquared:.4f}"])
    summary_data.append(['Adjusted R Square', f"{model.rsquared_adj:.4f}"])
    summary_data.append(['Standard Error', f"{model.bse.mean():.4f}"])
    summary_data.append(['Observations', f"{int(model.nobs)}"])
    summary_data.append([''])

    # Add ANOVA table
    summary_data.append(['ANOVA', ''])
    summary_data.append(['', 'df', 'SS', 'MS', 'F', 'Significance F'])
    for index, row in anova_table.iterrows():
        summary_data.append([str(index)] + [str(item) if item is not None else '' for item in row.tolist()])
    summary_data.append([''])

    # Add coefficients
    summary_data.append(['', 'Coefficients', 'Standard Error', 't Stat', 'P-value', 'Lower 95%', 'Upper 95%'])
    coeff_table = pd.read_html(model.summary().as_html(), header=0, index_col=0)[1].reset_index()

    # Separate 'Constant' and other variables
    constant_row = coeff_table[coeff_table.iloc[:, 0] == 'const'].iloc[0].tolist()
    x_vars = coeff_table[coeff_table.iloc[:, 0] != 'const'].iloc[:, 0].tolist()

    # Sort remaining x variables alphabetically
    x_vars_sorted = sorted(x_vars)

    # Add 'Constant' first
    summary_data.append([str(item) if item is not None else '' for item in constant_row])

    # Add sorted x variables
    for var in x_vars_sorted:
        row = coeff_table[coeff_table.iloc[:, 0] == var].iloc[0].tolist()
        summary_data.append([str(item) if item is not None else '' for item in row])

    # Convert summary data to DataFrame
    summary_df = pd.DataFrame(summary_data)

    # Convert DataFrame to dictionary
    data_dict = summary_df.to_dict(orient='index')

    # Correctly label columns in data_dict
    for key in data_dict:
        data_dict[key] = {str(i): str(value) for i, value in enumerate(data_dict[key].values())}

    # Create TableModel and TableCanvas
    table_model = TableModel()
    table_model.importDict(data_dict)
    table = TableCanvas(frame, model=table_model)
    table.show()


    def copy_to_clipboard():
        clipboard_data = '\n'.join(['\t'.join([str(cell) for cell in row]) for row in summary_data])
        results_window.clipboard_clear()
        results_window.clipboard_append(clipboard_data)
        messagebox.showinfo("Copied", "Data copied to clipboard")

    def export_to_excel():
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            summary_df.to_excel(file_path, index=False, header=False)
            messagebox.showinfo("Exported", f"Data exported to {file_path}")

    button_frame = tk.Frame(results_window)
    button_frame.pack(fill='x', pady=10)
    ttk.Button(button_frame, text="Copy to Clipboard", command=copy_to_clipboard).pack(side='left', padx=10)
    ttk.Button(button_frame, text="Export to Excel", command=export_to_excel).pack(side='left', padx=10)
    ttk.Button(button_frame, text="Close", command=results_window.destroy).pack(side='right', padx=10)


if __name__ == "__main__":
    read_xlsx_file()
