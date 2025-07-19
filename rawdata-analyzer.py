import pandas as pd

# Load Excel file
file_path = 'C:\Users\Laptop-SatyoYuwono\Downloads\rawdata.xls'
df = pd.read_excel(file_path)

# Keep only the columns you care about
columns_needed = [
    'miles_slc', 'hours_slc', 'minutes_slc',
    'date', 'type', 'engine_rpm', 'propeller_pitch',
    'me_hsfo_cons', 'me_lsfo_cons', 'ae_hsfo_cons', 'ae_lsfo_cons'
]
df = df[columns_needed].copy()

# Clean missing data (optional)
df = df.fillna(0)

# Add calculated columns
df['min_to_hrs'] = df['minutes_slc'] / 60
df['total_hrs'] = df['hours_slc'] + df['min_to_hrs']
df['vessel_speed'] = df['miles_slc'] / df['total_hrs'].replace(0, 1)
df['engine_distance'] = df['engine_rpm'] * df['total_hrs']
df['slip_percentage'] = ((df['engine_distance'] - df['miles_slc']) / df['engine_distance'].replace(0, 1)) * 100

# Save to new Excel
df.to_excel('processed_output.xlsx', index=False)

print("✅ File processed successfully. Output saved to 'processed_output.xlsx'")
