from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
from selenium.webdriver.common.action_chains import ActionChains
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os

#创建chrome对象，在电脑上打开一个窗口
browser = webdriver.Chrome()
browser.get("https://staging-eu01-vidaxl.demandware.net/on/demandware.store/Sites-Site/default/ViewApplication-DisplayWelcomePage")
sleep(2)
browser.set_window_size(1800,1000)
browser.implicitly_wait(50)
browser.find_element_by_id("idToken1").clear()
browser.find_element_by_id("idToken1").send_keys("mara.shu@habatrading.com")
browser.execute_script("arguments[0].click();", browser.find_element_by_id("loginButton_0"))
browser.implicitly_wait(50)
browser.find_element_by_id("idToken2").clear()
browser.find_element_by_id("idToken2").send_keys("699195Aaaa!")
browser.execute_script("arguments[0].click();", browser.find_element_by_id("loginButton_0"))
browser.implicitly_wait(50)

#进入正式页面
browser.find_element_by_link_text("Merchant Tools").click()
browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("vidaxl.at"))
browser.find_element_by_link_text("Products and Catalogs").click()
browser.find_element_by_link_text("Catalogs").click()
browser.find_element_by_link_text("vidaxl-catalog-webshop-eu-sku").click()
browser.implicitly_wait(50)

repeat = "0"

#获取表格数据
def obtain_table_data(url):
    # 获取表格数据.打开文件
    current_path = os.path.dirname(__file__)
    excel_paht = os.path.join(current_path,url) #r是转义保持原文件路径
    workbook = xlrd2.open_workbook(excel_paht)
    sheet=workbook.sheet_by_index(0)
    #双列表形式 ，一行一个用例
    all_case_info = []
    for i in range(1, sheet.nrows):  # 行数所以从1开始，0一般是名称之类的
        case_info = []
        for j in range(sheet.ncols):
            case_info.append(sheet.cell_value(i, j))
        all_case_info.append(case_info)  # 注意，python以对齐来确定循环的所定义区域
    return all_case_info

data_table = obtain_table_data("C:\\Users\\Mara.Shu\\Desktop\\sf_landing_page.xlsx")

#传入文件数据，添加landing page
# for num in range(0,len(data_table)):
#     repeat = "1"
#     browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("General"))
