import xlwings as xw
import pandas as pd
import numpy as np
import os

file_path = os.path.dirname(os.path.abspath(__file__))
print('当前工作的目录为：'+file_path)
app = xw.App(visible=True,add_book=False)
wb = app.books.open(file_path + r'\\01.xlsx')
sht = wb.sheets[0]
print('开始数据筛选')
sht.range('A1:A15').api.EntireRow.Delete() 
sht.range('F1:I1,K1').api.EntireColumn.Delete()
wb.save(file_path + r'\\用户时长排名.xlsx')
wb.close()

data = pd.read_excel(file_path + r'\\用户时长排名.xlsx')
group = ['/default','/226']
name = ['wifi','WIFI','无线路由器']
application = ['SSL数据:']
find_groupdata = data['组名'].str.contains('|'.join(group))
find_namedata = data['用户名'].str.contains('|'.join(name))
find_applicationdata = data['应用类型'].str.startswith('|'.join(application))
#print(data[find_groupdata])
#print(data[find_namedata])
#print(data[find_applicationdata])
df_group = pd.DataFrame(data[find_groupdata]+data[find_namedata]+data[find_applicationdata])
ls = []
ls.extend(df_group.index)
ls_array = np.array(ls)
final_ls = ls_array+2
str_ls = [str(i) for i in final_ls]
new_ls = ['A'+x for x in str_ls]
last_ls = str(','.join(new_ls))
#print(last_ls)
wb = app.books.open(file_path + r'\\用户时长排名.xlsx')
sht = wb.sheets[0]
sht.range(last_ls).api.EntireRow.Delete() 
sht.range('A12:A100').api.EntireRow.Delete() 
sht.range('A1:A12').api.EntireRow.Font.Name = '宋体'
sht.range('A1:A12').api.EntireRow.Font.Size = 11
wb.save(file_path + r'\\用户时长排名.xlsx')
wb.close()
app.quit()
print('数据筛选完成')