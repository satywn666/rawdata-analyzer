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

# 3. Convert time and calculate
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
# Filter: only total_hrs > 20
df = df[df['total_hrs'] > 20]

# 4. Show result
df.to_excel('C:/Users/Laptop-SatyoYuwono/Downloads/calculated_rawdata.xlsx', index=False)


