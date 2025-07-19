import pandas as pd

# Load Excel file
file_path = file_path = r'C:/Users/Laptop-SatyoYuwono/Downloads/rawdata.xls'
df = pd.read_excel(file_path, header=None)

# Keep only the columns you care about
columns_needed = [1, 2, 7, 8, 9, 15, 22, 44, 45, 48, 49]
df = df.iloc[:, columns_needed].copy()

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
