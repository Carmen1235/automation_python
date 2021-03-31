from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
from selenium.webdriver.common.action_chains import ActionChains
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os

#修改文件路径，确认时间区间,确认参数是否需要调整

# 获取表格数据
def obtain_table_data(url):
    # 获取表格数据.打开文件
    current_path = os.path.dirname(__file__)
    excel_paht = os.path.join(current_path, url)  # r是转义保持原文件路径
    workbook = xlrd2.open_workbook(excel_paht)
    sheet = workbook.sheet_by_index(0)
    # 双列表形式 ，一行一个用例

    all_case_info = []
    for i in range(1, sheet.nrows):  # 行数所以从1开始，0一般是名称之类的
        case_info = []
        for j in range(sheet.ncols):
            case_info.append(sheet.cell_value(i, j))
        all_case_info.append(case_info)  # 注意，python以对齐来确定循环的所定义区域
    return all_case_info

#创建chrome对象，在电脑上打开一个窗口
browser = webdriver.Chrome()
browser.get("https://admin.woger-cdn.com/index.php/admin/dashboard/index/key/ae57004afb8c09ab3969ca72b85962f7/")
browser.implicitly_wait(50)
#登录的部分
browser.set_window_size(1400,800)
browser.find_element_by_id("username").send_keys("Gloria.Wu@habatrading.com")
browser.find_element_by_id("login").send_keys("!gWJP!K8j*3%LG")
browser.find_element_by_class_name("form-button").click()
browser.refresh()
browser.implicitly_wait(50)


data_table = obtain_table_data(r"C:\Users\Mara.Shu\Desktop\sub-4. Head - Rattan outdoor funriture.xlsx")

link = browser.find_element_by_link_text("CMS") #定位到要悬停的元素
ActionChains(browser).move_to_element(link).perform()
Landingpage = browser.find_element_by_link_text("Image Sliders")
ActionChains(browser).move_to_element(Landingpage).perform()
browser.implicitly_wait(50) #属于隐式等待，5秒钟内只要找到了元素就开始执行，5秒钟后未找到，就超时
#点开需要悬停之后出现的元素
browser.find_element_by_link_text("List Sliders").click()
browser.implicitly_wait(50)

#这里主要是main的操作方法，虽然都差不多，但是需要需要对right-top，right-bottom和head-slider在进行更改
def by_object(country,name,block_id,image_file1,title1,image_url1,image_file2,title2,image_url2,image_file3,title3,image_url3,image_file4,title4,image_url4):
    browser.find_element_by_xpath('//*[@id="page:main-container"]/div[2]/table/tbody/tr/td[2]/button[@title="Add New"]').click()  # 点击添加
    browser.find_element_by_id("name").send_keys(name)
    browser.find_element_by_id("block_id").send_keys(block_id)
    store_id = browser.find_element_by_id("store")
    Select(store_id).deselect_all() #先清除选中的所有值
    Select(store_id).select_by_visible_text(country)
    browser.find_element_by_id("wis-save-and-continue").click()
    browser.implicitly_wait(50)
    #添加图片
    browser.find_element_by_id("imgslider_tabs_images").click()
    browser.find_element_by_id("wis-add-image").click()
    browser.find_element_by_id("image_file").send_keys(image_file1)
    browser.find_element_by_id("image_title").send_keys(title1)
    browser.find_element_by_id("sort_order").clear()
    browser.find_element_by_id("sort_order").send_keys("1")
    browser.find_element_by_id("image_url").send_keys(image_url1)
    browser.find_element_by_id("wis_imagesavebutton").click()
    sleep(5)
    browser.implicitly_wait(50)
    browser.execute_script("arguments[0].click();", browser.find_element_by_id("wis-add-image"))
    browser.find_element_by_id("image_file").send_keys(image_file2)
    browser.find_element_by_id("image_title").send_keys(title2)
    browser.find_element_by_id("sort_order").clear()
    browser.find_element_by_id("sort_order").send_keys("2")
    browser.find_element_by_id("image_url").send_keys(image_url2)
    browser.execute_script("arguments[0].click();", browser.find_element_by_id("wis_imagesavebutton"))
    sleep(5)
    browser.implicitly_wait(50)
    browser.execute_script("arguments[0].click();", browser.find_element_by_id("wis-add-image"))
    browser.find_element_by_id("image_file").send_keys(image_file3)
    browser.find_element_by_id("image_title").send_keys(title3)
    browser.find_element_by_id("sort_order").clear()
    browser.find_element_by_id("sort_order").send_keys("3")
    browser.find_element_by_id("image_url").send_keys(image_url3)
    browser.execute_script("arguments[0].click();", browser.find_element_by_id("wis_imagesavebutton"))
    sleep(5)
    browser.implicitly_wait(50)
    browser.execute_script("arguments[0].click();", browser.find_element_by_id("wis-add-image"))
    browser.find_element_by_id("image_file").send_keys(image_file4)
    browser.find_element_by_id("image_title").send_keys(title4)
    browser.find_element_by_id("sort_order").clear()
    browser.find_element_by_id("sort_order").send_keys("4")
    browser.find_element_by_id("image_url").send_keys(image_url4)
    browser.execute_script("arguments[0].click();", browser.find_element_by_id("wis_imagesavebutton"))
    sleep(3)
    browser.implicitly_wait(50)
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath('//*[@id="content"]/div/div[2]/p/button[@title="Save"]')) #保存
    browser.implicitly_wait(50)


for num in range(0, len(data_table)):
    by_object(data_table[num][1], data_table[num][2], data_table[num][3], data_table[num][4],data_table[num][5],data_table[num][6],data_table[num][7],data_table[num][8],data_table[num][9],data_table[num][10],data_table[num][11],data_table[num][12],data_table[num][13],data_table[num][14],data_table[num][15])
    browser.implicitly_wait(50)