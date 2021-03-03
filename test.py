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
browser.get("https://www.vidaxl.ee/auction/list")
sleep(2)
browser.set_window_size(1400,800)
browser.implicitly_wait(20)
browser.find_element_by_xpath("//input[@name='search']").send_keys("vidalx")
# #Next = driver.find_element_by_xpath("//input[@type='button' and @class='button']")
# browser.find_element_by_xpath("//img[@class='pos-absolute'][2]").click()

# #该方法用来确认元素是否存在，如果存在返回flag = true，否则返回false
# def isElementExist(element):
#     flag = True
#     try:
#         browser.find_element_by_xpath(element)
#         return flag
#
#     except:
#         flag = False
#         return flag
#
#
# flag = isElementExist("//img[@class='pos-absolute'][1]")
#
# if flag:
#     print("有弹框")
#
# else:
#     print("没有弹框")