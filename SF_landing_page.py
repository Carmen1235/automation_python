from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
from selenium.webdriver.common.action_chains import ActionChains
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os
from openpyxl import load_workbook
import re

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
browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='newCategory']")) #新建landing page

repeat = "0"

def create_page(country,category_id,name,description,page_title,seo_text):
    select_language = browser.find_element_by_xpath("//select[@name='LocaleId']")
    Select(select_language).select_by_value(country)
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_Id']").clear()
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_Id']").send_keys(category_id) #名字需要不同，不能一样
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_DisplayName']").clear()
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_DisplayName']").send_keys(name)
    browser.find_element_by_xpath("//textarea[@name='RegFormAddCategory_Description']").send_keys(description)
    if repeat == "0":
        print("11111111111111111111")
        browser.implicitly_wait(50)
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='create']"))
        browser.find_element_by_id("ValidFromDay").send_keys("02/20/2021")
        browser.find_element_by_id("ValidToDay").send_keys("03/08/2021")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.implicitly_wait(50)
        browser.find_element_by_link_text("Category Attributes").click()
        select_language = browser.find_element_by_xpath("//select[@name='LocaleId']")
        Select(select_language).select_by_value(country)
        browser.find_element_by_xpath("//input[@name='Meta4d0e39ba204e120610da62e2b4']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metacd1de64d47f33ab853a410d56a']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metad1fb3e7b9eafc17eb9f77e9f46']").send_keys("campaign-20-1-65-wasserspab")
        browser.find_element_by_xpath("//input[@name='Meta45d72a3f17c83b95a443fa57c1']").send_keys("FI_LP_Salesforce_1.png")
        browser.find_element_by_xpath("//input[@name='Meta2e671625926a8b1b0a4e40f419']").send_keys("rendering/campaignPage")
        browser.find_element_by_xpath("//textarea[@name='Meta64f25df909cb688e4d1998154a']").send_keys(seo_text)
        localized_online = browser.find_element_by_xpath("//select[@name='Meta5572cb622eba089436dbdc8665']")
        Select(localized_online).select_by_value("true")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.find_element_by_link_text("Search Refinement Definitions").click()
        browser.find_element_by_css_selector("a.action_link:first-child").click()  # 需要吧catalogs给block掉
        browser.implicitly_wait(50)
    else:
        print("2222222222222222222")
        browser.implicitly_wait(50)
        browser.find_element_by_id("ValidFromDay").clear()
        browser.find_element_by_id("ValidFromDay").send_keys("02/20/2021")
        browser.find_element_by_id("ValidToDay").clear()
        browser.find_element_by_id("ValidToDay").send_keys("03/08/2021")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.implicitly_wait(50)
        browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("Category Attributes"))
        select_language = browser.find_element_by_xpath("//select[@name='LocaleId']")
        Select(select_language).select_by_value(country)
        browser.find_element_by_xpath("//input[@name='Meta4d0e39ba204e120610da62e2b4']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metacd1de64d47f33ab853a410d56a']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metad1fb3e7b9eafc17eb9f77e9f46']").clear()
        browser.find_element_by_xpath("//input[@name='Metad1fb3e7b9eafc17eb9f77e9f46']").send_keys("campaign-20-1-65-wasserspab")
        browser.find_element_by_xpath("//input[@name='Meta45d72a3f17c83b95a443fa57c1']").clear()
        browser.find_element_by_xpath("//input[@name='Meta45d72a3f17c83b95a443fa57c1']").send_keys("FI_LP_Salesforce_1.png")
        browser.find_element_by_xpath("//input[@name='Meta2e671625926a8b1b0a4e40f419']").clear()
        browser.find_element_by_xpath("//input[@name='Meta2e671625926a8b1b0a4e40f419']").send_keys("rendering/campaignPage")
        browser.find_element_by_xpath("//textarea[@name='Meta64f25df909cb688e4d1998154a']").send_keys(seo_text)
        localized_online = browser.find_element_by_xpath("//select[@name='Meta5572cb622eba089436dbdc8665']")
        Select(localized_online).select_by_value("true")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.implicitly_wait(50)

def addproduct(): #添加产品
    browser.find_element_by_css_selector("span.icon-menu-menu_down_arrow:first-child").click()
    browser.find_element_by_xpath("//div[@title='Manage catalogs.']").click()
    browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("vidaxl-catalog-webshop-eu-sku"))
    browser.implicitly_wait(50)
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@value='All']"))
    browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("m_test_category1"))
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='assignProducts']"))
    browser.find_element_by_link_text("By ID").click()
    browser.find_element_by_xpath("//textarea[@name='WFSimpleSearch_IDList']").send_keys("8718475500636") #这里放添加产品的ean
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='findIDList']"))
    browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("Select All"))
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='assign']"))
    browser.implicitly_wait(50)

def rebuild(): #选择checkbox
    browser.find_element_by_link_text("Merchant Tools").click()
    browser.find_element_by_link_text("Search").click()
    browser.find_element_by_link_text("Search Indexes").click()
    browser.find_element_by_id("idxgrp_prd").click()
    browser.find_element_by_id("idxgrp_suggest").click()
    browser.find_element_by_id("idxgrp_activedata").click()
    browser.find_element_by_id("sharedidxgrp_availability").click()
    browser.find_element_by_xpath("//button[@name='index']").click()

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
for num in range(0,len(data_table)):
    create_page(data_table[num][1],data_table[num][2],data_table[num][3],data_table[num][4],data_table[num][5],data_table[num][6])
    repeat = "1"
    browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("General"))

addproduct()
rebuild()