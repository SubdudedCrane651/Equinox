import xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Open the existing .xlsm file
file_path = r'C:\\Users\\rchrd\\Documents\\Richard\\Electricity for Equinox.xlsm'
wb = xw.Book(file_path)

# Select the sheet you want to work with (replace 'Sheet1' with your actual sheet name)
sheet = wb.sheets['Sheet1']

# Read data from C9 to C100
data_range = sheet.range('C9:C1000').value

# Get the current date
today = datetime.now()
start_date = datetime(2024, 10, 1)  # Starting from October 2024

# Calculate monthly expenditures
monthly_expenditures = []
current_expenditure = 0
months = []

for expenditure in data_range:
    if expenditure is None:
        if current_expenditure > 0:
            months.append(start_date.strftime("%B %Y"))
            monthly_expenditures.append(current_expenditure)
            start_date += relativedelta(months=1)
            current_expenditure = 0  # Reset for the next month
    else:
        current_expenditure += expenditure

# Add the last month's expenditure if there's any remaining
if current_expenditure != 0:
    months.append(start_date.strftime("%B %Y"))
    monthly_expenditures.append(current_expenditure)

# Prepare data with months and expenditures
data = [['Month', 'Expenditure (CAD)']]
for month, expenditure in zip(months, monthly_expenditures):
    data.append([month, expenditure])

# Write the updated data to starting at K16
sheet.range('K16:L' + str(16 + len(data) - 1)).value = data

# Create the bar chart at N16
chart = sheet.charts.add(left=sheet.range('N16').left, top=sheet.range('N16').top)
chart.chart_type = 'bar_clustered'
chart.set_source_data(sheet.range('K16:L' + str(16 + len(data) - 1)))

# Set chart title and axis labels
chart.api[1].ChartTitle.Text = 'Electric Car Expenditures per Month (CAD)'
chart.api[1].Axes(1).HasTitle = True
chart.api[1].Axes(1).AxisTitle.Text = 'Month'
chart.api[1].Axes(2).HasTitle = True
chart.api[1].Axes(2).AxisTitle.Text = 'Expenditure (CAD)'

# Save and close the workbook
#wb.save()
#wb.close()
