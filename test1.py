from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
from selenium.webdriver.common.action_chains import ActionChains
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os
from openpyxl import load_workbook
import re
from selenium.webdriver.common.keys import Keys

#创建chrome对象，在电脑上打开一个窗口
# browser = webdriver.Chrome()
# browser.get("https://woger.login.myclang.com/clang/build/production/ClangUI/index.html")
# browser.set_window_size(1400,800)
# browser.implicitly_wait(50)
# browser.find_element_by_id("textfield-1015-inputEl").send_keys("mara.shu@HabaTrading.com")
# browser.find_element_by_id("textfield-1016-inputEl").send_keys("L2cLE44Ry1VX")
# browser.execute_script("arguments[0].click();", browser.find_element_by_id("button-1020-btnInnerEl"))
# browser.implicitly_wait(300)
# browser.find_element_by_xpath("//li[@data-recordid='152']").click()
text = "a"
if text == "a":
    print("bfsdj")
else:
    print("hfsdjkh")