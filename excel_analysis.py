import pandas as pd
import openpyxl


#-------------------------------------------------------------------------------------------
# reading data from excel file
data_file = 'Sales Records.xlsx'

workbook = openpyxl.load_workbook('data//'+data_file)
sheets = workbook.sheetnames

# print excel workbook and sheets that are in the workbook
print(f'> excel file is: ',data_file)
print(f'> excel sheet names: ',sheets)

#-------------------------------------------------------------------------------------------
# get rows and columns for data in 'americas' sheet
americas_sheet = workbook['americas']
americas_data = americas_sheet.values
americas_df = pd.DataFrame(americas_data)
print(americas_df)  

header = americas_df.iloc[0]
americas_df = americas_df[1:]
americas_df.columns = header
print(americas_df)

americas_shape = americas_df.shape
print(f'> rows: ',americas_shape[0])
print(f'> columns: ',americas_shape[1])


#-------------------------------------------------------------------------------------------
# let's do some calculations
# get grand total of revenue for america
america_rev_total = americas_df['Total Revenue'].sum()
print(f'> total revenue = {america_rev_total}')

# get grand total of cost for america
america_cost_total = americas_df['Total Cost'].sum()
print(f'> total cost = {america_cost_total}')

# get grand total of cost for america
america_profit_total = americas_df['Total Profit'].sum()
print(f'> total profit = {america_profit_total}')

america_profit_total_check = america_rev_total - america_cost_total
print(f'> total profit (check) = {america_profit_total_check}')


#-------------------------------------------------------------------------------------------
# add grand total of revenue for america and output to excel workbook
americas_df.loc['Total', 'Total Revenue']= americas_df['Total Revenue'].sum()
#print(sales_data)

americas_df.loc['Total', 'Total Cost']= americas_df['Total Cost'].sum()
americas_df.loc['Total', 'Total Profit']= americas_df['Total Profit'].sum()
americas_df.loc['Total', 'Units Sold']= americas_df['Units Sold'].sum()
#print(sales_data)


americas_df['Region'].iloc[-1] = americas_df.index[-1]
#print(americas_df)

'''
# save updated dataframe to excel file
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "americas"

for r in dataframe_to_rows(americas_df, index=False, header=True):
    ws.append(r)
 
wb.save('data//Sales Records processed.xlsx')
'''


#-------------------------------------------------------------------------------------------
# add average and median for america and output to excel workbook
americas_df.loc['Median', 'Total Revenue']= americas_df['Total Revenue'].median()

americas_df.loc['Median', 'Total Cost']= americas_df['Total Cost'].median()
americas_df.loc['Median', 'Total Profit']= americas_df['Total Profit'].median()
americas_df.loc['Median', 'Units Sold']= americas_df['Units Sold'].median()
americas_df.loc['Median', 'Unit Price']= americas_df['Unit Price'].median()
americas_df.loc['Median', 'Unit Cost']= americas_df['Unit Cost'].median()

americas_df['Region'].iloc[-1] = americas_df.index[-1]

americas_df.loc['Mean', 'Total Revenue']= americas_df['Total Revenue'].mean()


americas_df.loc['Mean', 'Total Cost']= americas_df['Total Cost'].mean()
americas_df.loc['Mean', 'Total Profit']= americas_df['Total Profit'].mean()
americas_df.loc['Mean', 'Units Sold']= americas_df['Units Sold'].mean()
americas_df.loc['Mean', 'Unit Price']= americas_df['Unit Price'].mean()
americas_df.loc['Mean', 'Unit Cost']= americas_df['Unit Cost'].mean()

americas_df['Region'].iloc[-1] = americas_df.index[-1]


# save updated dataframe to excel file
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "americas"

for r in dataframe_to_rows(americas_df, index=False, header=True):
    ws.append(r)
 
#wb.save('data//Sales Records processed.xlsx')


#-------------------------------------------------------------------------------------------
# format the header
font_black_bold = openpyxl.styles.Font(color='000000', bold=True)
bd_thick = openpyxl.styles.Side(style='thick', color="000000")

for cell in ws[1]:
    cell.font = font_black_bold
    cell.border = openpyxl.styles.Border(bottom=bd_thick)
    
ws.freeze_panes = ws['A2']
wb.save('data//Sales Records processed.xlsx')