import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Series, Reference
pd.set_option("display.max_columns", 500)

wb = Workbook()
#grab the active worksheet
ws = wb.create_sheet("Analysis")
del wb['Sheet']

raw_df = pd.read_csv("all_delays.csv")

station_analysis = raw_df.groupby(["Station"])['Day'].count().reset_index()
station_analysis.columns = ["Station", "CountOfDelays"]
station_analysis = station_analysis.sort_values(by=["CountOfDelays"], ascending=False)


for r in dataframe_to_rows(station_analysis, index=False, header=True):
    ws.append(r)


chart1 = BarChart()
chart1.type = "col"
chart1.style = 10
chart1.title = "Bar Chart"
chart1.y_axis.title = 'NumberOfDelays'
chart1.x_axis.title = 'Station Name'

data = Reference(ws, min_col=2, min_row=1, max_row=6, max_col=2)
cats = Reference(ws, min_col=1, min_row=2, max_row=6)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.shape = 4
ws.add_chart(chart1, "D2")


wb.save("pandas_openpyxl_practice.xlsx")
