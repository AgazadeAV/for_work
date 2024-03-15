import pandas as pd
import matplotlib.pyplot as plt
from calendar import month_name
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

# Sample data: Replace this with your actual data
data = {
    'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    'Min_Rate': [10, 12, 15, 11, 13, 14, 16, 18, 20, 19, 17, 16],
    'Max_Rate': [20, 22, 25, 21, 23, 24, 26, 28, 30, 29, 27, 26]
}

df = pd.DataFrame(data)

# Create Excel workbook and add data to it
wb = Workbook()
ws = wb.active
ws.append(['Month', 'Min_Rate', 'Max_Rate'])
for row in df.values:
    ws.append(row)

# Create bar chart
chart = BarChart()
chart.title = "Min-Max Rates by Months"
chart.x_axis.title = "Month"
chart.y_axis.title = "Rate"

data = Reference(ws, min_col=2, min_row=1, max_row=len(df)+1, max_col=3)
categories = Reference(ws, min_col=1, min_row=2, max_row=len(df)+1)

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

# Add chart to worksheet
ws.add_chart(chart, "D2")

# Save the workbook
wb.save("min_max_rates_calendar.xlsx")
