import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

# Assuming your data is in a CSV file (you can replace this with your actual data source)
full_chart_data = pd.read_csv('your_file.csv')  # Update this with your actual file

# Ensure 'date' column is in datetime format
full_chart_data['date'] = pd.to_datetime(full_chart_data['date'])

# Setup plot
fig, ax = plt.subplots(figsize=(10, 6))

# Colors for each metric
colors = ['b', 'g', 'r', 'c', 'm']

# Set the bar width
bar_width = 0.1

# Iterate through the metrics you want to plot (replace 'metric1', 'metric2', etc., with your actual metrics)
metrics = ['metric1', 'metric2', 'metric3']  # Replace with actual metric column names

# Loop through metrics and plot each as a bar chart
for i, metric in enumerate(metrics):
    ax.bar(full_chart_data['date'] + pd.Timedelta(i, 'D'), full_chart_data[metric], 
           width=bar_width, label=metric, color=colors[i])

# Set labels and title
ax.set_xlabel('Date')
ax.set_ylabel('Value')
ax.set_title('Your Chart Title')

# Add legend
ax.legend()

# Show the plot
plt.show()

# Working with OpenPyXL (if needed for exporting to Excel or dealing with labels)
# Open an Excel file
wb = openpyxl.load_workbook('your_excel_file.xlsx')  # Replace with your actual Excel file
ws = wb.active

# Let's assume you want to add data labels in an Excel chart:
chart = openpyxl.chart.BarChart()
data = openpyxl.chart.Reference(ws, min_col=2, min_row=2, max_col=3, max_row=10)  # Adjust range accordingly
categories = openpyxl.chart.Reference(ws, min_col=1, min_row=2, max_row=10)  # Adjust range accordingly
chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

# Add chart to worksheet
ws.add_chart(chart, "E5")

# Optionally, adjust labels or other settings if needed
# For example, to add a data label to each point:
# This part assumes you're using the openpyxl.chart package properly with labels and titles
# For openpyxl's DataLabel functionality, you'd normally use chart.dataLabels attributes

# Save the modified workbook
wb.save('modified_excel_file.xlsx')
