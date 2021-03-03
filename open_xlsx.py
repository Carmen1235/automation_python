import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os

def obtain_table_data(url):
    # 获取表格数据.打开文件
    current_path = os.path.dirname(__file__)
    excel_paht = os.path.join(current_path,url) #r是转义保持原文件路径
    workbook = xlrd2.open_workbook(excel_paht)

    sheet=workbook.sheet_by_index(0)

    #双列表形式 ，一行一个用例
    all_case_info = []
    for i in range(1,sheet.nrows): #行数所以从1开始，0一般是名称之类的
        case_info=[]
        for j in range(sheet.ncols):
            case_info.append(sheet.cell_value(i,j))
        all_case_info.append(case_info)   #注意，python以对齐来确定循环的所定义区域
    return all_case_info

data_table = obtain_table_data(r"C:\Users\Mara.Shu\Desktop\week06\webshop link.xlsx")
# print(data_table[0][1])
for num in range(0,len(data_table)):
    print(data_table[num][0])
    print(data_table[num][1])
    print(data_table[num][2])
    print(data_table[num][3])
    print(data_table[num][4])