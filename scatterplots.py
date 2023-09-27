# -*- coding: utf-8 -*-
"""
Created on Tue Jul 18 14:22:15 2023

@author: zndr2
"""
# import pandas as pd
# import matplotlib.pyplot as plt
# import numpy as np
# from scipy import stats

# def create_scatterplot(x_values, y_values, column_name, plot_title):
#     # Plot scatter plot
#     plt.scatter(x_values, y_values, label='Data')
#     # linear regression
#     slope, intercept, _, _, _ = stats.linregress(x_values, y_values)
#     # correlation coefficient (r value)
#     correlation_coefficient = np.corrcoef(x_values, y_values)[0, 1]
#     # trend line
#     plt.plot(x_values, slope * x_values + intercept, color='r', label='Trendline')
#     # Add r to plot
#     plt.text(
#         x_values.mean(),
#         y_values.min(),
#         f'r = {correlation_coefficient:.2f}',
#         horizontalalignment='center',
#         verticalalignment='bottom'
#     )
#     # title, axis labels, and legend
#     plt.title(f"Scatterplot - {plot_title}")
#     plt.xlabel("X")
#     plt.ylabel("Y")
#     plt.legend()
#     plt.show()
    
    
# # paired t test function 

# def plot_excel(excel_file, sheet_name, x_column, y_columns, start_row, end_row):
#     # Read the Excel file into a DataFrame
#     df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
#     # Extract x values (column E)
#     x_values = df.iloc[start_row-1:end_row-1, x_column-1].astype(float)

#     # Extract y values for each column
#     for col in y_columns:
#         y_values = df.iloc[start_row-1:end_row-1, col-1].astype(float)
#         column_name = df.columns[col-1]
#         plot_title = df.iloc[2, col-1]  # Values from row 4
#         create_scatterplot(x_values, y_values, column_name, plot_title)

# # Excel file, sheet name, column ranges
# excel_file = 'Z:/Treadway Pilot/Ketamine Pilot/Regional Analysis/combined_results_FITT2_results_SI_all_k9_anatomic.xlsx'
# sheet_name = 'PSS_correlation_no_outliers'
# x_column = 5
# y_columns = range(8, 28)
# start_row = 4
# end_row = 22

# # Call function
# plot_excel(excel_file, sheet_name, x_column, y_columns, start_row, end_row)























#exclude values?

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats

def create_scatterplot(x_values, y_values, column_name, plot_title):
    # Filter out data points with missing y values
    valid_data_mask = ~np.isnan(y_values)
    x_values_filtered = x_values[valid_data_mask]   
    y_values_filtered = y_values[valid_data_mask]

    # Plot scatter plot
    plt.scatter(x_values_filtered, y_values_filtered, label='Data')
    # linear regression
    slope, intercept, _, _, _ = stats.linregress(x_values_filtered, y_values_filtered)
    # correlation coefficient (r value)
    correlation_coefficient = np.corrcoef(x_values_filtered, y_values_filtered)[0, 1]
    # trend line
    plt.plot(x_values_filtered, slope * x_values_filtered + intercept, color='r', label='Trendline')
    # Add r to plot
    plt.text(
        x_values_filtered.mean(),
        y_values_filtered.min(),
        f'r = {correlation_coefficient:.2f}',
        horizontalalignment='center',
        verticalalignment='bottom'
    )
    # title, axis labels, and legend
    plt.title(f"Scatterplot - {plot_title}")
    plt.xlabel("X")
    plt.ylabel("Y")
    plt.legend()
    plt.show()
    
    
# paired t test function 

def plot_excel(excel_file, sheet_name, x_column, y_columns, start_row, end_row):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Extract x values (column E)
    x_values = df.iloc[start_row-1:end_row-1, x_column-1].astype(float)

    # Extract y values for each column
    for col in y_columns:
        y_values = df.iloc[start_row-1:end_row-1, col-1].astype(float)
        column_name = df.columns[col-1]
        plot_title = df.iloc[2, col-1]  # Values from row 4
        create_scatterplot(x_values, y_values, column_name, plot_title)

# Excel file, sheet name, column ranges
excel_file = 'Z:/Treadway Pilot/Ketamine Pilot/Regional Analysis/combined_results_FITT2_results_SI_all_k9_anatomic.xlsx'
sheet_name = 'PSS_correlation_no_outliers'
x_column = 5
y_columns = range(8, 28)
start_row = 4
end_row = 22

# Call function
plot_excel(excel_file, sheet_name, x_column, y_columns, start_row, end_row)

