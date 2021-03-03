import requests
import bs4

#让访问的网页知道这是一个正常访问的浏览器
headers={}
#获取网页
page_obj=requests.get("https://www.vidaxl.nl/",headers=headers)
#解析网页
bas_obj=bs4.BeautifulSoup(page_obj.text,"lxml")
list_obj=bas_obj.find_all("div",attrs={"class":"layout"})
print(list_obj)