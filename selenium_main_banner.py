from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
from selenium.webdriver.common.action_chains import ActionChains
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os

#注意事项，需要输入搜索的值来进行查询
#修改查询的是main，right-top，right-bottom和head-slider
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


data_table = obtain_table_data(r"C:\Users\Mara.Shu\Desktop\Headslider.xlsx")

#查询magento的值可以输入main，right-top，right-bottom和head-slider
def main(search_value):
    link = browser.find_element_by_link_text("CMS") #定位到要悬停的元素
    ActionChains(browser).move_to_element(link).perform()
    Landingpage = browser.find_element_by_link_text("Image Sliders")
    ActionChains(browser).move_to_element(Landingpage).perform()
    browser.implicitly_wait(50) #属于隐式等待，5秒钟内只要找到了元素就开始执行，5秒钟后未找到，就超时
    #点开需要悬停之后出现的元素
    browser.find_element_by_link_text("List Sliders").click()
    browser.implicitly_wait(50)
    browser.find_element_by_id("imgslider_filter_block_id").send_keys(search_value)
    browser.find_element_by_css_selector("td.a-right button.task").click() #搜索为main的的国家
    num = browser.find_element_by_css_selector("td.pager select") #选择数量为200的进行显示
    Select(num).select_by_value("200")
    browser.implicitly_wait(50)

#这里主要是main的操作方法，虽然都差不多，但是需要需要对right-top，right-bottom和head-slider在进行更改
def by_object(country,image_file,image_title,banner_title,banner_subtitle,button_text,image_url,sort_order):
    #根据文本定位点击的国家
    tr = browser.find_elements_by_css_selector("table#imgslider_table tbody tr")
    for len_tr in range(0,len(tr)):
        text = tr[len_tr].get_attribute("textContent")  # 获取元素标签的内容textContent 获取元素内的全部HTML innerHTML 获取包含选中元素的HTML outerHTML
        if text.find(country) != -1 :
            tr[len_tr].click()
            break
    browser.find_element_by_id("imgslider_tabs_images").click()
    browser.implicitly_wait(50)
    browser.find_element_by_id("wis-add-image").click() #点击image按钮并点击添加按钮
    browser.implicitly_wait(50)
    browser.find_element_by_id("image_file").send_keys(image_file) #添加图片文件路径
    browser.find_element_by_id("image_title").send_keys(image_title) #添加image_title
    browser.find_element_by_id("image_from").send_keys("03/30/21") #添加date from
    browser.find_element_by_id("image_to").send_keys("05/02/21") #添加date to
    browser.find_element_by_id("sort_order").clear()
    browser.find_element_by_id("sort_order").send_keys(int(sort_order)) #添加sort_order
    Size_type = browser.find_element_by_id("size_type") #选择size——type
    Select(Size_type).select_by_value("xl")
    browser.find_element_by_id("banner_title").send_keys(banner_title)
    browser.find_element_by_id("banner_subtitle").send_keys(banner_subtitle)
    browser.find_element_by_id("button_text").send_keys(button_text)
    browser.find_element_by_id("image_url").send_keys(image_url)
    browser.find_element_by_id("wis_imagesavebutton").click()
    browser.implicitly_wait(50)

main("head-slider")
for num in range(0, len(data_table)):
    by_object(data_table[num][1], data_table[num][2], data_table[num][3], data_table[num][4],data_table[num][5],data_table[num][6],data_table[num][7],data_table[num][8])
    browser.implicitly_wait(50)
    element1 = browser.find_element_by_css_selector("button.back:first-child")
    browser.execute_script("arguments[0].click();", element1)