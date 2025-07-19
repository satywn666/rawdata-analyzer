import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------------------- Load Excel File ----------------------
file_path = r'C:\Users\Laptop-SatyoYuwono\Downloads\FZ-JAN-JUN2.xls'
df = pd.read_excel(file_path, header=None)

# Convert date column (index 1) to datetime
df[1] = pd.to_datetime(df[1], format="%Y/%m/%d", errors='coerce')

# ---------------------- Filter by Date Range ----------------------
voyage_start_date = pd.to_datetime('2025-04-01')
voyage_end_date = pd.to_datetime('2025-04-21')
df_voyage = df[(df[1] >= voyage_start_date) & (df[1] <= voyage_end_date)]

# ---------------------- FO Consumption ----------------------
fo_rob_initial = df_voyage.iloc[0, 4]  # Column E
fo_rob_final = df_voyage.iloc[-1, 4]   # Column E
supplied_fo = df_voyage[34].sum()     # Column AI
fo_consumed = fo_rob_initial - fo_rob_final
if fo_consumed < 0:
    fo_consumed += supplied_fo

# ---------------------- Prepare Data ----------------------
columns_needed = [1, 2, 7, 8, 9, 11, 15, 22, 44, 45, 48, 49]
df_calc = df_voyage.iloc[:, columns_needed].copy()
df_calc.columns = [
    'date', 'type', 'miles_slc', 'hours_slc', 'minutes_slc', 'wind_force',
    'engine_rpm', 'propeller_pitch', 'me_hsfo_cons', 'me_lsfo_cons',
    'ae_hsfo_cons', 'ae_lsfo_cons'
]

# ---------------------- Calculations ----------------------
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
df_calc['performance_speed'] = ((
    df_calc['engine_rpm'] * 
    df_calc['propeller_pitch'] * 
    (1 - df_calc['slip_percentage'] / df_calc['engine_distance']) * 
    60
) / 1852
)


# ---------------------- Beaufort Mapping ----------------------
def get_beaufort_speed_loss(wind_force):
    if wind_force <= 3:
        return 0.0
    elif wind_force <= 6:
        return 0.0
    elif wind_force <= 10:
        return 0.015
    elif wind_force <= 16:
        return 0.03
    elif wind_force <= 21:
        return 0.05
    elif wind_force <= 27:
        return 0.075
    elif wind_force <= 33:
        return 0.10
    elif wind_force <= 40:
        return 0.13
    elif wind_force <= 47:
        return 0.17
    elif wind_force <= 55:
        return 0.22
    else:
        return 0.25

df_calc['speed_loss_pct'] = df_calc['wind_force'].apply(get_beaufort_speed_loss)
df_calc['weather_factor'] = df_calc['performance_speed'] * df_calc['speed_loss_pct']

# ---------------------- Current Factor ----------------------
df_calc['current_factor'] = df_calc['vessel_speed'] - df_calc['performance_speed'] + df_calc['weather_factor']

# ---------------------- Filtered / Unfiltered ----------------------
df_filtered = df_calc[df_calc['total_hrs'] > 20]
df_unfiltered = df_calc.copy()

# ---------------------- Average ----------------------
avg_subset = df_unfiltered[df_unfiltered['total_hrs'] > 10]
avg_data = {
    'date': 'AVERAGE (>10hrs)', 'type': '', 'miles_slc': '', 'hours_slc': '', 'minutes_slc': '',
    'wind_force': '', 'engine_rpm': '', 'propeller_pitch': '', 'me_hsfo_cons': avg_subset['me_hsfo_cons'].mean(),
    'me_lsfo_cons': avg_subset['me_lsfo_cons'].mean(), 'ae_hsfo_cons': avg_subset['ae_hsfo_cons'].mean(),
    'ae_lsfo_cons': avg_subset['ae_lsfo_cons'].mean(), 'min_to_hrs': '', 'total_hrs': '',
    'vessel_speed': avg_subset['vessel_speed'].mean(), 'engine_distance': '', 'slip_percentage': avg_subset['slip_percentage'].mean(),
    'performance_speed': avg_subset['performance_speed'].mean(), 'speed_loss_pct': '',
    'weather_factor': avg_subset['weather_factor'].mean(), 'current_factor': avg_subset['current_factor'].mean()
}
df_unfiltered = pd.concat([df_unfiltered, pd.DataFrame([avg_data])], ignore_index=True)

# ---------------------- Save to Excel ----------------------
output_path = 'C:/Users/Laptop-SatyoYuwono/Downloads/FZ-WEATHER-RESULT-6_1.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_filtered.to_excel(writer, sheet_name='Filtered Sailing', index=False)
    df_unfiltered.to_excel(writer, sheet_name='Unfiltered Sailing', index=False)

    # Summary Sheet
    summary_data = pd.DataFrame({
        'Description': ['FO ROB Initial', 'FO ROB Final', 'Supplied FO', 'Total FO Consumed'],
        'Value': [fo_rob_initial, fo_rob_final, supplied_fo, fo_consumed]
    })
    summary_data.to_excel(writer, sheet_name='FO Consumption Summary', index=False)

    # Chart Sheet
    df_chart = df_calc[df_calc['total_hrs'] > 10][[
        'date', 'engine_rpm', 'me_hsfo_cons', 'vessel_speed', 'slip_percentage',
        'weather_factor', 'current_factor', 'performance_speed'
    ]].copy()
    charter_speed = 13.5
    df_chart['charter_speed'] = charter_speed

    chart_sheet = writer.book.create_sheet('Performance Chart')
    for r in dataframe_to_rows(df_chart, index=False, header=True):
        chart_sheet.append(r)

    # Chart 1 - Bar Chart
    bar_chart = BarChart()
    bar_chart.title = "Performance Metrics"
    data = Reference(chart_sheet, min_col=2, min_row=1, max_col=5, max_row=len(df_chart)+1)
    cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=len(df_chart)+1)
    bar_chart.add_data(data, titles_from_data=True)
    bar_chart.set_categories(cats)
    chart_sheet.add_chart(bar_chart, 'K2')


print("Done: FO consumption, voyage data, performance & weather impact saved.")
