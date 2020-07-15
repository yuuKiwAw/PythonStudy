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
sht.range('A1,D1:H1,J1').api.EntireColumn.Delete()
wb.save(file_path + r'\\用户流量排名.xlsx')
wb.close()

data = pd.read_excel(file_path + r'\\用户流量排名.xlsx')
group = ['/default','/226']
name = ['wifi','WIFI','无线路由器']
find_groupdata = data['组名'].str.contains('|'.join(group))
find_namedata = data['用户名'].str.contains('|'.join(name))
df = pd.DataFrame(data[find_groupdata]+data[find_namedata])
ls = []
ls.extend(df.index)
ls_array = np.array(ls)
final_ls = ls_array+2
str_ls = [str(i) for i in final_ls]
new_ls = ['A'+x for x in str_ls]
last_ls = str(','.join(new_ls))
#print(ls_array)
wb = app.books.open(file_path + r'\\用户流量排名.xlsx')
sht = wb.sheets[0]
sht.range(last_ls).api.EntireRow.Delete() 
sht.range('A12:A100').api.EntireRow.Delete() 
sht.range('A1:A12').api.EntireRow.Font.Name = '宋体'
sht.range('A1:A12').api.EntireRow.Font.Size = 12
wb.save(file_path + r'\\用户流量排名.xlsx')
wb.close()

wb = app.books.open(file_path + r'\\用户流量排名.xlsx')
excel_path = (file_path + r'\\用户流量排名.xlsx')
formate_data = pd.read_excel(excel_path,sheet_name='sheet1',usecols=[2])
fd_formate_data = pd.DataFrame(formate_data)
ls_formate_data = fd_formate_data['总流量'].values.tolist()
for i in range(len(ls_formate_data)):
    ls_formate_data[i] = ls_formate_data[i].replace('Gb','')
#print(ls_formate_data)
wb = app.books.open(file_path + r'\\用户流量排名.xlsx')
sht = wb.sheets[0]
sht.range('C2:C11').options(transpose=True).value = ls_formate_data
sht.api.Columns(1).Insert()
sht.api.Columns(3).Copy(sht.api.Columns(1))
sht.range('C1').api.EntireColumn.Delete()
wb.save(file_path + r'\\用户流量排名.xlsx')
wb.close()

app.quit()
