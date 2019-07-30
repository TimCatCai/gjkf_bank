import pandas as pd
import  xlrd
import numpy
input_data_path = r"C:\Users\TimCat\Desktop\jin_nian.xlsx"
xls_data = pd.DataFrame(pd.read_excel(input_data_path))
ids = list(xls_data['身份证号'])
print(ids)
amounts = xls_data['合同余额']
if '440903199207061515' in ids:
    print(amounts[ids.index('440903199207061515')])
