import xlwings as xw
import pandas as pd
import numpy as np
import os

file_path = os.path.dirname(os.path.abspath(__file__))
print('当前工作的目录为：'+file_path)
app = xw.App(visible=True,add_book=False)
wb = app.books.open(file_path + r'\\应用流量排行.xlsx')
sht = wb.sheets[0]
sht.range('A1:A17').api.EntireRow.Delete()
sht.range('C1:G1,I1').api.EntireColumn.Delete()
wb.save(file_path + r'\\部门流量排名.xlsx')
wb.close()

data = pd.read_excel(file_path + r'\\部门流量排名.xlsx')
group = ['/default','/226','所有组']
find_groupdata = data['组名'].str.contains('|'.join(group))
df = pd.DataFrame(data[find_groupdata])

ls = []
ls.extend(df.index)
ls_array = np.array(ls)
final_ls = ls_array+2
str_ls = [str(i) for i in final_ls]
new_ls = ['A'+x for x in str_ls]
last_ls = str(','.join(new_ls))

wb = app.books.open(file_path + r'\\部门流量排名.xlsx')
sht = wb.sheets[0]
sht.range(last_ls).api.EntireRow.Delete() 
wb.save(file_path + r'\\部门流量排名.xlsx')
wb.close()

excel_path = (file_path + r'\\部门流量排名.xlsx')
formate_data = pd.read_excel(excel_path,sheet_name='sheet1',usecols=[2])
fd_formate_data = pd.DataFrame(formate_data)
ls_formate_data = fd_formate_data['总流量'].values.tolist()
for i in range(len(ls_formate_data)):
    ls_formate_data[i] = ls_formate_data[i].replace('Gb','')
#print(ls_formate_data)

wb = app.books.open(file_path + r'\\部门流量排名.xlsx')
sht = wb.sheets[0]
sht.range('C2').options(transpose=True).value = ls_formate_data
sht.range('A1:A40').api.EntireRow.Font.Name = '宋体'
sht.range('A1:A40').api.EntireRow.Font.Size = 11
wb.save(file_path + r'\\部门流量排名.xlsx')
wb.close()

app.quit()