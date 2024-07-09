import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
pd.set_option("display.max_columns", 500)

wb = Workbook()
# grab the active worksheet
ws = wb.create_sheet("Analysis")
del wb['Sheet']

raw_df = pd.read_csv("all_delays.csv")

station_analysis = raw_df.groupby(["Station", "Code"])['Day'].count().reset_index()
station_analysis.columns = ["Station", "Code", "CountOfDelays"]
station_analysis = station_analysis.sort_values(by=["CountOfDelays"], ascending=False)


for r in dataframe_to_rows(station_analysis, index=False, header=True):
    ws.append(r)

wb.save("pandas_openpyxl_practice.xlsx")

