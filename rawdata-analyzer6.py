import pandas as pd

# 1. Load Excel
file_path = r'C:\Users\Laptop-SatyoYuwono\Downloads\rawdata.xls'
df = pd.read_excel(file_path, header=None)

# 2. Convert date column (index 1) to datetime
df[1] = pd.to_datetime(df[1], errors='coerce')

# 3. Manually input date range
voyage_start_date = pd.to_datetime('2025-01-08')
voyage_end_date = pd.to_datetime('2025-01-25')

# 4. Filter by date range (inclusive)
df_voyage = df[(df[1] >= voyage_start_date) & (df[1] <= voyage_end_date)]

# 5. FO Consumption Calculation
fo_rob_initial = df_voyage.iloc[0, 4]  # Column E
fo_rob_final = df_voyage.iloc[-1, 4]   # Column E
supplied_fo = df_voyage[34].sum()      # Column AI
fo_consumed = fo_rob_initial - fo_rob_final
if fo_consumed < 0:
    fo_consumed += supplied_fo

# 6. Select and rename needed columns for calculation
columns_needed = [1, 2, 7, 8, 9, 15, 22, 44, 45, 48, 49]
df_calc = df_voyage.iloc[:, columns_needed].copy()
df_calc.columns = [
    'date', 'type', 'miles_slc', 'hours_slc', 'minutes_slc',
    'engine_rpm', 'propeller_pitch', 'me_hsfo_cons', 'me_lsfo_cons',
    'ae_hsfo_cons', 'ae_lsfo_cons'
]

# 7. Calculate time and performance
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

# 8. Filtered and unfiltered data
df_filtered = df_calc[df_calc['total_hrs'] > 20]
df_unfiltered = df_calc.copy()

# 9. Calculate average values for total_hrs > 10
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

# 10. Save to Excel
output_path = 'C:/Users/Laptop-SatyoYuwono/Downloads/voyage_rawdata_separated.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_filtered.to_excel(writer, sheet_name='Filtered Sailing', index=False)
    df_unfiltered.to_excel(writer, sheet_name='Unfiltered Sailing', index=False)

    # FO Consumption Summary
    summary_data = pd.DataFrame({
        'Description': ['FO ROB Initial', 'FO ROB Final', 'Supplied FO', 'Total FO Consumed'],
        'Value': [fo_rob_initial, fo_rob_final, supplied_fo, fo_consumed]
    })
    summary_data.to_excel(writer, sheet_name='FO Consumption Summary', index=False)

print("Done: FO consumption and voyage data saved with summaries.")
