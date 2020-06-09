import pandas as pd 
import xlwings as xl 
import seaborn as sns

# get this workbook
wb = xl.Book(r'data.xlsx')

# getting the sheets references
dataSheet = wb.sheets['data']
plotSheet = wb.sheets['graph']

# transforming data in pandas dataframe
data_range_str = dataSheet.used_range.get_address(row_absolute=False, column_absolute=False)
data_range = dataSheet.range(data_range_str).value
dt = pd.DataFrame(data_range[1:], columns=data_range[0])

#filtering data
site_param = plotSheet.range('B1').value
country_param = plotSheet.range('B2').value
product_param = plotSheet.range('B3').value



#ploting graph



