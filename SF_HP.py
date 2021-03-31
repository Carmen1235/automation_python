from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select #下拉框的选择操作
from selenium.webdriver.common.action_chains import ActionChains
import xlrd2 #xlrd只支持xls，xlrd2可以支持xlsx格式
import os
from selenium.webdriver.common.keys import Keys

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

data_table = obtain_table_data("C:\\Users\\Mara.Shu\\Desktop\\sf_home_page.xlsx")

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
merchant_tools = browser.find_element_by_css_selector("span.icon-menu-menu_down_arrow:first-child")
merchant_tools.click()
browser.find_element_by_link_text("Content Assets").click()
browser.implicitly_wait(50)
browser.find_element_by_xpath("//button[@name='new']").click()  #新建一个页面
browser.find_element_by_xpath("//input[@name='Meta15e6961db65d149cf67b1f2310']").send_keys("week-01-t") #添加id
online = browser.find_element_by_id("Meta599c32498805d799f92d82b7a4")
Searchable = browser.find_element_by_id("Meta5c3368fb60e6c8a617d5ca66c2")
Select(online).select_by_value("true")
Select(Searchable).select_by_value("true")
browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='apply']"))  #建立好了一个id页面，现在需要在里面添加国家
browser.implicitly_wait(50)

# 正式添加国家的数据
def main(county,show,cgid,id,url,title,sub,button):
    select_country = browser.find_element_by_xpath("//select[@name='LocaleId']")
    Select(select_country).select_by_value(county)
    browser.find_element_by_xpath("//textarea[@name='Meta87c098e11cb2d8af5383d83905']").send_keys('<style>@media only screen and (min-width:780px) {'
        +'.main-sub-title {'
        +'font-size: 22px;'
        +'}'
        +'}'
        +'</style>'
        +'<div class="mr-lg-2 flex-grow-1">'
        +'<a href="$url('+show+', '+cgid+', '+id+')$" class="text-decoration-none">'
        +'<div class="d-flex flex-column align-items-center justify-content-between py-2 py-sm-4 px-3 rounded-sm product-banner product-banner-main" style="background-image: url('+url+');">'
        +'<div>'
        +'<h2 class="h1 text-center text-white text-shadow">'+title+'</h2>'
        +'<p class="font-weight-bold text-center text-white mb-0 text-shadow title main-sub-title">'+sub+'</p>'
        +'</div>'
        +'<p><button class="btn btn-primary">'+button+'</button></p>'
        +'</div>'
        +'</div>'
        +'</a>')
    browser.execute_script("arguments[0].click();", browser.find_element_by_xpath("//button[@name='apply']"))
    browser.implicitly_wait(50)

# #传入文件数据，添加landing page
for num in range(0,len(data_table)):
    main(data_table[num][0],data_table[num][2],data_table[num][3],data_table[num][4],data_table[num][5],data_table[num][6],data_table[num][7],data_table[num][8])

browser.find_element_by_css_selector("span.icon-menu-menu_down_arrow:first-child").click()
browser.find_element_by_link_text("Content Slots").click()
search_input = browser.find_element_by_id("searchTriggerField")
search_input.send_keys("homepage")
search_input.send_keys(Keys.ENTER)
browser.find_element_by_id("ext-gen256").click()
browser.find_element_by_link_text("homepage-banners").click()
browser.implicitly_wait(50)
#根据文本定位点击的国家
tr = browser.find_elements_by_css_selector("div.x-grid3-row")
for len_tr in range(0,len(tr)):
    text = tr[len_tr].get_attribute("textContent")  # 获取元素标签的内容textContent 获取元素内的全部HTML innerHTML 获取包含选中元素的HTML outerHTML
    if text.find("Winter sale flash sale: hardware% 1.14") != -1 :  #安排时间选择哪个可以直接在这边进行更改
        tr[len_tr].click()
        break
browser.find_element_by_id("id").clear()
browser.find_element_by_id("id").send_keys("id")
browser.find_element_by_id("fieldEnabled").click()
browser.find_element_by_css_selector("div.x-combo-list-item:first-child").click() #选择yes
browser.find_element_by_id("description").clear()
browser.find_element_by_id("description").send_keys("id_name")
browser.find_element_by_id("content").clear()
browser.find_element_by_id("content").send_keys("content")

#安排时间
browser.find_elements_by_css_selector("td.x-grid3-td-schedule")[1].click()
browser.find_element_by_id("slotScheduleTrigger").send_keys("03/05/2021 12:00 am - 03/19/2021 12:15 am")

