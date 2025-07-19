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

# 3. Convert 'date' column to datetime format and handle errors
df['date'] = pd.to_datetime(df['date'], errors='coerce')

# 4. Convert time and calculate
df['min_to_hrs'] = df['minutes_slc'] / 60
df['total_hrs'] = df['hours_slc'] + df['min_to_hrs']

# Avoid divide by zero
df['vessel_speed'] = df.apply(lambda row: row['miles_slc'] / row['total_hrs'] if row['total_hrs'] > 0 else 0, axis=1)

# Engine distance: RPM * pitch (meters) * total_hrs * 60 (min/hr), convert to NM (1 NM = 1852 m)
df['engine_distance'] = df.apply(
    lambda row: (row['engine_rpm'] * row['propeller_pitch'] * row['total_hrs'] * 60) / 1852 if row['total_hrs'] > 0 else 0,
    axis=1
)

# Slip percentage
df['slip_percentage'] = df.apply(
    lambda row: (1 - row['miles_slc'] / row['engine_distance']) * 100 if row['engine_distance'] > 0 else 0,
    axis=1
)

# 5. Manual input for voyage start and end dates (format: 'YYYY-MM-DD')
voyage_start_date = '2025-01-08'  # Replace with the actual start date you want
voyage_end_date = '2025-01-26'  # Replace with the actual end date you want + 1 day

# 6. Convert start and end dates to datetime
voyage_start_date = pd.to_datetime(voyage_start_date)
voyage_end_date = pd.to_datetime(voyage_end_date)

# 7. Filter data based on the input start and end dates (inclusive)
df_voyage = df[(df['date'] >= voyage_start_date) & (df['date'] <= voyage_end_date)]

# 8. Create two DataFrames:
#    - Filtered (only total_hrs > 20)
#    - Unfiltered (all data within the date range, no total_hrs filter)

df_filtered = df_voyage[df_voyage['total_hrs'] > 20]  # Filtered by total_hrs > 20
df_unfiltered = df_voyage  # Unfiltered by total_hrs, includes all data in the date range

# 9. Create a new Excel file with two sheets:
#    - One for the filtered data (total_hrs > 20)
#    - One for the unfiltered data (all data within the date range)

# Define output Excel path
output_path = 'C:/Users/Laptop-SatyoYuwono/Downloads/voyage_rawdata_separated.xlsx'

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # Write filtered data to the first sheet
    df_filtered.to_excel(writer, sheet_name='Filtered Sailing', index=False)
    
    # Write unfiltered data (all data within the date range) to the second sheet
    df_unfiltered.to_excel(writer, sheet_name='Unfiltered Sailing', index=False)

print(f"Data has been written to: {output_path}")
