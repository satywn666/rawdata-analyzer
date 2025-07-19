import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------------------- Load Data ----------------------
file_path = r'C:\Users\Laptop-SatyoYuwono\Downloads\FZ-JAN-JUN2.xls'
df = pd.read_excel(file_path, header=None)
df[1] = pd.to_datetime(df[1], format="%Y/%m/%d", errors='coerce')

voyage_start_date = pd.to_datetime('2025-04-01')
voyage_end_date = pd.to_datetime('2025-04-21')
df_voyage = df[(df[1] >= voyage_start_date) & (df[1] <= voyage_end_date)]

fo_rob_initial = df_voyage.iloc[0, 4]
fo_rob_final = df_voyage.iloc[-1, 4]
supplied_fo = df_voyage[34].sum()
fo_consumed = fo_rob_initial - fo_rob_final
if fo_consumed < 0:
    fo_consumed += supplied_fo

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
df_calc['performance_speed'] = (
    df_calc['engine_rpm'] * df_calc['propeller_pitch'] * (1 - df_calc['slip_percentage'] / 100) * 60
) / 1852

def get_beaufort_speed_loss(wind_force):
    if wind_force <= 3: return 0.0
    elif wind_force <= 6: return 0.0
    elif wind_force <= 10: return 0.015
    elif wind_force <= 16: return 0.03
    elif wind_force <= 21: return 0.05
    elif wind_force <= 27: return 0.075
    elif wind_force <= 33: return 0.10
    elif wind_force <= 40: return 0.13
    elif wind_force <= 47: return 0.17
    elif wind_force <= 55: return 0.22
    else: return 0.25

df_calc['speed_loss_pct'] = df_calc['wind_force'].apply(get_beaufort_speed_loss)
df_calc['weather_factor'] = df_calc['performance_speed'] * df_calc['speed_loss_pct']
df_calc['current_factor'] = df_calc['vessel_speed'] - df_calc['performance_speed'] + df_calc['weather_factor']

# ---------------------- Chart Data Preparation ----------------------
df_chart = df_calc[df_calc['total_hrs'] > 10][[
    'date', 'weather_factor', 'current_factor', 'performance_speed'
]].copy()
df_chart['charter_speed'] = 13.5

# ---------------------- Write to Excel ----------------------
output_path = 'C:/Users/Laptop-SatyoYuwono/Downloads/FZ-STACKED-SIGIT-ANJING.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_chart.to_excel(writer, sheet_name='Chart Data', index=False)
    chart_sheet = writer.book['Chart Data']

    # ---------------------- Create Chart ----------------------
    bar_chart = BarChart()
    bar_chart.type = "col"
    bar_chart.grouping = "stacked"
    bar_chart.overlap = 100
    bar_chart.title = "Weather + Current Factor with Speed Overlay"
    bar_chart.y_axis.title = "Knots"
    bar_chart.x_axis.title = "Date"

    # Stack: Weather + Current
    bar_data = Reference(chart_sheet, min_col=2, min_row=1, max_col=3, max_row=len(df_chart) + 1)
    bar_cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=len(df_chart) + 1)
    bar_chart.add_data(bar_data, titles_from_data=True)
    bar_chart.set_categories(bar_cats)

    # Overlay Line: Performance + Charter Speed
    for col in range(4, 6):
        line = Series(Reference(chart_sheet, min_col=col, min_row=1, max_row=len(df_chart)+1),
                      title=chart_sheet.cell(row=1, column=col).value)
        line.chart_type = "line"
        bar_chart.series.append(line)

    chart_sheet.add_chart(bar_chart, "H2")

print("✅ Done. Chart exported with correct stacking and line overlays.")
