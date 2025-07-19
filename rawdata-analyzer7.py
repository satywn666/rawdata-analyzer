import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

# Locating & Load Excel File <-----------------------------
file_path = r'C:\Users\Laptop-SatyoYuwono\Downloads\Jan-Juli-MF.xls'
df = pd.read_excel(file_path, header=None)

# Convert date column (index 1) to datetime
df[1] = pd.to_datetime(df[1], errors='coerce')

# Manually input date range <------------------------------
voyage_start_date = pd.to_datetime('2025-01-1')
voyage_end_date = pd.to_datetime('2025-01-31') # add one day

# Filter by date range (inclusive)
df_voyage = df[(df[1] >= voyage_start_date) & (df[1] <= voyage_end_date)]

# FO Consumption Calculation
fo_rob_initial = df_voyage.iloc[0, 4]  # Column E
fo_rob_final = df_voyage.iloc[-1, 4]   # Column E
supplied_fo = df_voyage[34].sum()      # Column AI
fo_consumed = fo_rob_initial - fo_rob_final
if fo_consumed < 0:
    fo_consumed += supplied_fo

# Select and rename needed columns for calculation
columns_needed = [1, 2, 7, 8, 9, 15, 22, 44, 45, 48, 49]
df_calc = df_voyage.iloc[:, columns_needed].copy()
df_calc.columns = [
    'date', 'type', 'miles_slc', 'hours_slc', 'minutes_slc',
    'engine_rpm', 'propeller_pitch', 'me_hsfo_cons', 'me_lsfo_cons',
    'ae_hsfo_cons', 'ae_lsfo_cons'
]

# Calculate time and performance
df_calc['min_to_hrs'] = df_calc['minutes_slc'] / 60
df_calc['total_hrs'] = df_calc['hours_slc'] + df_calc['min_to_hrs']
df_calc['vessel_speed'] = df_calc.apply(lambda row: row['miles_slc'] / row['total_hrs'] if row['total_hrs'] > 0 else 0, axis=1)
df_calc['engine_distance'] = df_calc.apply(
    lambda row: (row['engine_rpm'] * row['propeller_pitch'] * row['total_hrs'] * 60) / 1852 if row['total_hrs'] > 0 else 0,
    axis=1
)
df_calc['slip_percentage'] = df_calc.apply(
    lambda row: (1 - row['miles_slc'] / row['engine_distance']) * 100 if row['engine_distance'] > 0 else 0,
    axis=1
)

# Filtered and unfiltered data
df_filtered = df_calc[df_calc['total_hrs'] > 20]
df_unfiltered = df_calc.copy()

# Calculate average values for total_hrs > 10
avg_subset = df_unfiltered[df_unfiltered['total_hrs'] > 10]
avg_data = {
    'date': 'AVERAGE (>10hrs)',
    'type': '',
    'miles_slc': '',
    'hours_slc': '',
    'minutes_slc': '',
    'engine_rpm': '',
    'propeller_pitch': '',
    'me_hsfo_cons': avg_subset['me_hsfo_cons'].mean(),
    'me_lsfo_cons': avg_subset['me_lsfo_cons'].mean(),
    'ae_hsfo_cons': avg_subset['ae_hsfo_cons'].mean(),
    'ae_lsfo_cons': avg_subset['ae_lsfo_cons'].mean(),
    'min_to_hrs': '',
    'total_hrs': '',
    'vessel_speed': avg_subset['vessel_speed'].mean(),
    'engine_distance': '',
    'slip_percentage': ''
}
df_unfiltered = pd.concat([df_unfiltered, pd.DataFrame([avg_data])], ignore_index=True)

# Save to Excel with chart <-------------------------------
output_path = 'C:/Users/Laptop-SatyoYuwono/Downloads/MF-JAN.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_filtered.to_excel(writer, sheet_name='Filtered Sailing', index=False)
    df_unfiltered.to_excel(writer, sheet_name='Unfiltered Sailing', index=False)

    # FO Consumption Summary
    summary_data = pd.DataFrame({
        'Description': ['FO ROB Initial', 'FO ROB Final', 'Supplied FO', 'Total FO Consumed'],
        'Value': [fo_rob_initial, fo_rob_final, supplied_fo, fo_consumed]
    })
    summary_data.to_excel(writer, sheet_name='FO Consumption Summary', index=False)

    # Chart Data Preparation
    df_chart = df_calc[df_calc['total_hrs'] > 10][['date', 'engine_rpm', 'me_hsfo_cons', 'vessel_speed', 'slip_percentage']]
    df_chart.sort_values(by='date', inplace=True)

    # Write data for chart
    df_chart_sheet = writer.book.create_sheet('Performance Chart')
    for r in dataframe_to_rows(df_chart, index=False, header=True):
        df_chart_sheet.append(r)

    # Create Bar Chart
    chart = BarChart()
    chart.type = "col"
    chart.title = "Vessel Performance Metrics"
    chart.y_axis.title = "Values"
    chart.x_axis.title = "Date"
    chart.style = 10
    chart.width = 30
    chart.height = 12
    chart.grouping = "clustered"
    chart.overlap = 0

    # Created by Satyo Yuwono

    # Data range for chart (excluding 'date' column)
    data = Reference(df_chart_sheet, min_col=2, min_row=1, max_col=5, max_row=len(df_chart) + 1)
    cats = Reference(df_chart_sheet, min_col=1, min_row=2, max_row=len(df_chart) + 1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    df_chart_sheet.add_chart(chart, "G2")

print("Done: FO consumption, voyage data, and performance chart saved.")
