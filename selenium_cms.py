from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
from selenium.webdriver.common.action_chains import ActionChains
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os

#注意事项
# 1、注意修改时间
# 2、判断条件增加一个判断条件


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
#定位到要悬停的元素
Promotions = browser.find_element_by_link_text("Promotions")
ActionChains(browser).move_to_element(Promotions).perform()
browser.find_element_by_link_text("Shopping Cart Price Rules").click()
browser.implicitly_wait(50) #属于隐式等待，5秒钟内只要找到了元素就开始执行，5秒钟后未找到，就超时

def by_value(name,description,country,coupon_code,discount_amount):
    browser.find_element_by_css_selector("button.add:first-child").click()
    browser.find_element_by_id("rule_name").send_keys("name") #rule_name
    browser.find_element_by_id("rule_description").send_keys("description") #description
    websites = browser.find_element_by_id("rule_website_ids")
    Select(websites).select_by_visible_text("Austria") #选择国家
    rule_coupon_type = browser.find_element_by_id("rule_coupon_type")
    Select(rule_coupon_type).select_by_visible_text("Specific Coupon") #coupon_type
    browser.find_element_by_id("rule_coupon_code").send_keys("ddd") #coupon_code
    browser.find_element_by_id("rule_uses_per_customer").clear()
    browser.find_element_by_id("rule_uses_per_customer").send_keys("0") #uses_per_customer
    browser.find_element_by_id("rule_sort_order").clear()
    browser.find_element_by_id("rule_sort_order").send_keys("0") #Priority
    browser.find_element_by_id("rule_from_date").send_keys("02/17/2021")
    browser.find_element_by_id("rule_to_date").send_keys("02/18/2021")
    browser.implicitly_wait(50)

    element1 = browser.find_element_by_id("promo_catalog_edit_tabs_conditions_section")
    browser.execute_script("arguments[0].click();", element1) #点击Conditions的页面
    browser.find_element_by_css_selector("ul#conditions__1__children li span.rule-param-new-child").click()
    add_conditions = browser.find_element_by_id("conditions__1__new_child")
    Select(add_conditions).select_by_visible_text("Subtotal") #选择条件
    browser.find_element_by_link_text("...").click()
    browser.find_element_by_id("conditions__1--1__value").send_keys("149") #添加值
    browser.implicitly_wait(50)

    element2 = browser.find_element_by_id("promo_catalog_edit_tabs_actions_section")
    browser.execute_script("arguments[0].click();", element2) #点击auciton页面
    apply = browser.find_element_by_id("rule_simple_action")
    Select(apply).select_by_visible_text("Fixed amount discount for whole cart")
    browser.find_element_by_id("rule_discount_amount").clear()
    browser.find_element_by_id("rule_discount_amount").send_keys("15") #discount_amount

    browser.find_element_by_css_selector("ul#actions__1__children li span.rule-param-new-child").click()
    add_conditions1 = browser.find_element_by_id("actions__1__new_child")
    Select(add_conditions1).select_by_visible_text("Category")
    browser.find_element_by_link_text("...").click()
    browser.find_element_by_id("actions__1--1__value").send_keys("vidaxl")

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

data_table = obtain_table_data(r"C:\Users\Mara.Shu\Desktop\mian_banner.xlsx")