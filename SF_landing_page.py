from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os

#注意事项
#日期，图片，表格的名字，产品的eam

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
browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("vidaxl-catalog-webshop-eu-sku"))
browser.implicitly_wait(50)
browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='newCategory']")) #新建landing page

repeat = "0"

def create_page(country,category_id,name,page_title,page_url,seo_text,standard_image):
    select_language = browser.find_element_by_xpath("//select[@name='LocaleId']")
    Select(select_language).select_by_value(country)
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_Id']").clear()
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_Id']").send_keys(category_id) #名字需要不同，不能一样
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_DisplayName']").clear()
    browser.find_element_by_xpath("//input[@name='RegFormAddCategory_DisplayName']").send_keys(name)
    if repeat == "0":
        browser.implicitly_wait(50)
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='create']"))
        browser.find_element_by_id("ValidFromDay").send_keys("03/29/2021")
        browser.find_element_by_id("ValidToDay").send_keys("01/01/2022")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.implicitly_wait(50)
        browser.find_element_by_link_text("Category Attributes").click()
        select_language = browser.find_element_by_xpath("//select[@name='LocaleId']")
        Select(select_language).select_by_value(country)
        browser.find_element_by_xpath("//input[@name='Meta4d0e39ba204e120610da62e2b4']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metacd1de64d47f33ab853a410d56a']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metad1fb3e7b9eafc17eb9f77e9f46']").send_keys(page_url) #Page url
        browser.find_element_by_xpath("//input[@name='Meta45d72a3f17c83b95a443fa57c1']").send_keys(standard_image)
        browser.find_element_by_xpath("//input[@name='Meta2e671625926a8b1b0a4e40f419']").send_keys("rendering/campaignPage")
        browser.find_element_by_xpath("//textarea[@name='Meta64f25df909cb688e4d1998154a']").send_keys(seo_text)
        localized_online = browser.find_element_by_xpath("//select[@name='Meta5572cb622eba089436dbdc8665']")
        Select(localized_online).select_by_value("true")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.find_element_by_link_text("Search Refinement Definitions").click()
        browser.find_element_by_css_selector("a.action_link:first-child").click()  # 需要吧catalogs给block掉
        browser.implicitly_wait(50)
    else:
        browser.implicitly_wait(50)
        browser.find_element_by_id("ValidFromDay").clear()
        browser.find_element_by_id("ValidFromDay").send_keys("03/29/2021")
        browser.find_element_by_id("ValidToDay").clear()
        browser.find_element_by_id("ValidToDay").send_keys("01/01/2022")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.implicitly_wait(50)
        browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("Category Attributes"))
        select_language = browser.find_element_by_xpath("//select[@name='LocaleId']")
        Select(select_language).select_by_value(country)
        browser.find_element_by_xpath("//input[@name='Meta4d0e39ba204e120610da62e2b4']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metacd1de64d47f33ab853a410d56a']").send_keys(page_title)
        browser.find_element_by_xpath("//input[@name='Metad1fb3e7b9eafc17eb9f77e9f46']").clear()
        browser.find_element_by_xpath("//input[@name='Metad1fb3e7b9eafc17eb9f77e9f46']").send_keys(page_url)
        browser.find_element_by_xpath("//input[@name='Meta45d72a3f17c83b95a443fa57c1']").clear()
        browser.find_element_by_xpath("//input[@name='Meta45d72a3f17c83b95a443fa57c1']").send_keys(standard_image)
        browser.find_element_by_xpath("//input[@name='Meta2e671625926a8b1b0a4e40f419']").clear()
        browser.find_element_by_xpath("//input[@name='Meta2e671625926a8b1b0a4e40f419']").send_keys("rendering/campaignPage")
        browser.find_element_by_xpath("//textarea[@name='Meta64f25df909cb688e4d1998154a']").send_keys(seo_text)
        localized_online = browser.find_element_by_xpath("//select[@name='Meta5572cb622eba089436dbdc8665']")
        Select(localized_online).select_by_value("true")
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='update']"))
        browser.implicitly_wait(50)

def addproduct(category_id): #添加产品
    browser.find_element_by_css_selector("span.icon-menu-menu_down_arrow:first-child").click()
    browser.find_element_by_xpath("//div[@title='Manage catalogs.']").click()
    browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("vidaxl-catalog-webshop-eu-sku"))
    browser.implicitly_wait(50)
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@value='All']"))
    browser.execute_script("arguments[0].click();", browser.find_element_by_link_text(category_id))
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='assignProducts']"))
    browser.find_element_by_link_text("By ID").click()
    browser.find_element_by_xpath("//textarea[@name='WFSimpleSearch_IDList']").send_keys("ean=8718475506973,8718475501046,8718475501039,8719883732077,8719883732053,8719883732183,8718475614845,8718475607014,8719883732121,8718475621812,8718475504917,8719883732190,8718475851332,8718475614838,8718475504900,8719883732237,8718475502722,8718475851349,8719883732220,8718475960324,8718475607007,8718475960317,8718475502715,8718475502500,8719883732244,8719883732213,8719883732206,8719883860411,8719883796659,8719883868196,8718475609643,8719883760445,8718475606994,8719883732176,4251682255592,7081452051935,8718475901785,8719128193045,5053163481198,7439621595504,8718475901907,8718475901761,8718475616535,8718475963356,8718475907480,8718475963318,8718475616498,8718475901884,8719883725307,8718475901792,8718475901877,8718475601647,8718475992264,4051814367687,8718475621836,8718475500902,8718475702245,8718475612551,8719883832364,8719883832487,8718475607687,8719883727264,8719883727196,8719883835921,8719883727349,8719883727257,8719883851792,8719883754895,8719883727202,8719883726052,8719883725277,8719883754901,8719883746043,8719883743769,8719883754888,8719883835969,8719883832517,8719883835976,8719883743783,8719883743677,8719883836003,8719883754871,8719883832357,8719883835907,8718475601531,8719883731315,8719883784878,8719883836010,8719883813851,8719883743752,8719883890463,8719883741079,8719883835952,8719883785066,8719883784175,8719883726113,8718475504351,8718475500810,8718475504368,8718475504931,8718475506348,8718475609629,8719883729329,8718475607816,8718475601661,8718475621867,8718475501640,8719883725154,8718475697282,8718475504740,8719883729312,8719883725192,8718475607793,8719883725239,8718475505358,8719883867700,8718475607823,8718475607250,8719883731353,8719883729299,8719883729336,8718475617099,8718475501657,8719883725208,8719883796680,8718475581611,8719883725178,8719883725161,8718475697251,8719883743806,8718475501664,8719883755038,8719883726106,8718475503590,8718475705109,8719883727622,8719883725215,8719883743820,8719883784892,8719883755250,8719883726274,8719883725130,8718475732600,8719883784168,8718475607649,8719883732329,8719883552569,8719883785059,8719883784144,8719883726397,8719883732336,8719883726403,8719883732343,8719883785097,8719883784182,8719883751702,8719883871721,8719883726328,8718475501695,8718475601593,8718475501688,8718475743132,8718475743125,8718475607021,8718475743118,8718475601708,8719883800790,8718475607038,8719883727370,8718475601692,8718475743156,8718475743163,8719883552590,8719883552606,8719883765464,8718475743149,8719883727226,8719883727387,8719883727219,8719883765457,8719883552613,8719883755342,8719883755397,8719883755359,8719883755328,8719883755434,8719883729503,8719883784984,8719883784977,8719883729510,8719883755335,8719883784991,8719883729497") #这里放添加产品的ean
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='findIDList']")) #查找find
    # 获取元素标签的内容textContent 获取元素内的全部HTML innerHTML 获取包含选中元素的HTML outerHTML
    text = browser.find_element_by_xpath("//button[@value='All']").get_attribute("textContent")
    if text.find("All") != -1:
        #如果有超过多的sku则先显示全部
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@value='All']"))
        browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("Select All"))
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='assign']"))
        browser.implicitly_wait(50)
    else:
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

data_table = obtain_table_data("C:\\Users\\Mara.Shu\\Desktop\\4. Head - Rattan outdoor funriture.xlsx")

#传入文件数据，添加landing page
for num in range(0,len(data_table)):
    create_page(data_table[num][1],data_table[num][2],data_table[num][3],data_table[num][4],data_table[num][5],data_table[num][6],data_table[num][7])
    repeat = "1"
    browser.execute_script("arguments[0].click();", browser.find_element_by_link_text("General"))

addproduct(data_table[0][2])
rebuild()