import pandas as pd

# 1. Load Excel
file_path = r'C:\Users\Laptop-SatyoYuwono\Downloads\rawdata.xls'
df = pd.read_excel(file_path, header=None)

# 2. Select and rename columns
columns_needed = [1, 2, 7, 8, 9, 15, 22, 44, 45, 48, 49]
df = df.iloc[:, columns_needed].copy()
df.columns = [
    'date', 'type', 'miles_slc', 'hours_slc', 'minutes_slc',
    'engine_rpm', 'propeller_pitch', 'me_hsfo_cons', 'me_lsfo_cons',
    'ae_hsfo_cons', 'ae_lsfo_cons'
]

# 3. Convert date column to datetime
df['date'] = pd.to_datetime(df['date'], errors='coerce')

# 4. Calculate time and performance
df['min_to_hrs'] = df['minutes_slc'] / 60
df['total_hrs'] = df['hours_slc'] + df['min_to_hrs']
df['vessel_speed'] = df.apply(lambda row: row['miles_slc'] / row['total_hrs'] if row['total_hrs'] > 0 else 0, axis=1)
df['engine_distance'] = df.apply(
    lambda row: (row['engine_rpm'] * row['propeller_pitch'] * row['total_hrs'] * 60) / 1852 if row['total_hrs'] > 0 else 0,
    axis=1
)
df['slip_percentage'] = df.apply(
    lambda row: (1 - row['miles_slc'] / row['engine_distance']) * 100 if row['engine_distance'] > 0 else 0,
    axis=1
)

# 5. Manually input date range
voyage_start_date = pd.to_datetime('2025-01-08')
voyage_end_date = pd.to_datetime('2025-01-25')

# 6. Filter by date range (inclusive)
df_voyage = df[(df['date'] >= voyage_start_date) & (df['date'] <= voyage_end_date)]

# 7. Filtered and unfiltered DataFrames
df_filtered = df_voyage[df_voyage['total_hrs'] > 20]
df_unfiltered = df_voyage.copy()

# 8. Calculate average for selected columns where total_hrs > 10
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

# 9. Write to Excel
output_path = 'C:/Users/Laptop-SatyoYuwono/Downloads/voyage_rawdata_separated.xlsx'

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_filtered.to_excel(writer, sheet_name='Filtered Sailing', index=False)
    df_unfiltered.to_excel(writer, sheet_name='Unfiltered Sailing', index=False)

print("Finished: Added average for selected columns in Unfiltered Sailing sheet.")
