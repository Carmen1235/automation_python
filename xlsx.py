import xlsxwriter
import pandas as pd
#输入数据 分别代表每一列的数据
df=pd.DataFrame({'name':['meat','rice'],'price':[12,3],'quantity':[10,100]})
#创建xlsx文件。
workbook=xlsxwriter.Workbook('products.xlsx')
#新增工作簿。
worksheet=workbook.add_worksheet()
#字体加粗，蓝色背景，紫色背景
format_columname=workbook.add_format({'bold':True,'font_color':'blue','bg_color':'purple'})
#设置价格的数值格式，并添加下划线
format_price=workbook.add_format({'num_format': '$#,##0','underline':True})
#设置字体
format_products=workbook.add_format({'font_name':'Times New Roman'})
#worksheet.write函数写入第一行列名，参数分别表示行、列、数据、数据格式
for col in range(len(df.columns)):
    worksheet.write(0, col, df.columns[col], format_columname)
#分别写入元素。
for row in range(2):
    worksheet.write(row + 1, 0, df.name[row])
for row in range(2):
    worksheet.write(row + 1, 1, df.price[row], format_price)
for row in range(2):
    worksheet.write(row + 1, 2, df.quantity[row], format_products)
workbook.close()