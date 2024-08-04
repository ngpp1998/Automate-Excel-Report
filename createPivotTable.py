# pip install openpyxl
import pandas as pd

df = pd.read_excel('supermarket_sales.xlsx')

df = df[['Gender', 'Product line', 'Total']]

pivotTable = df.pivot_table(index='Gender', columns='Product line', values='Total', aggfunc='sum').round(0)

pivotTable.to_excel('pivot_table.xlsx', 'Report', startrow=4)