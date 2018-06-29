# -*- coding: utf-8 -*-
from selenium import webdriver
from time import sleep
from openpyxl import workbook ,load_workbook ,Workbook
import os ,time ,ytFuntion

def sheet_date():
    sheet["F" + str(i)].value = time.strftime("%y_%m_%d") #檢查日期
    sheet["G" + str(i)].value = time.strftime("%H_%M_%S") #檢查時間
    
chrome_path = "D:\selenium_driver_chrome\chromedriver.exe" #webdriver放置資料夾
mobileEmulation = {'deviceName': 'iPhone 6/7/8'}
options = webdriver.ChromeOptions()
options.add_experimental_option('mobileEmulation', mobileEmulation)

webDriver = webdriver.Chrome(chrome_path ,chrome_options=options)
test_web = ytFuntion.test_web(webDriver)

wb = load_workbook("檢查JS用_H5.xlsx")
sheet = wb["web"] # 獲取一張表

testWebUrl = input("請輸入測試站點的url(Ex.http://m.winvip66.acgtest.com):")
sheet["A1"].value = testWebUrl
webDriver.get(str(testWebUrl) + str(sheet["D2"].value).strip())
testSiteID = input("請輸入測試SiteID:")
sheet["A2"].value = testSiteID
j = 0
for i in range(8,13):#找siteID所在位置
    try:
        int(test_web["site.config"][test_web["site.config"].index("siteId") + i])
        j += 1
    except:
        pass    
newSiteConfig = test_web["site.config"][:test_web["site.config"].index("siteId") + 8] + str(testSiteID) + test_web["site.config"][test_web["site.config"].index("siteId") + 8 + j:]
if len(testSiteID.strip()) == 0:
    pass
else:
    test_web["site.config"] = newSiteConfig #更換siteConfig
textCheck = input("請輸入要檢測的字:")
sheet["A3"].value = textCheck
print("檢測目標:" + textCheck)
print()

for i in range(2 ,len(sheet["B"])+1):
    if i == 10:
        input("請輸入帳密,再按enter。")
        '''test_web.elementClick("亲，请登录",3)
        test_web.elementSendKeys("input[tag=帐号]" ,6 ,text = "ytau1")
        test_web.elementSendKeys("input[tag=密码]" ,6 ,text = "qwe123")
        test_web.elementClick("[class='mainColorBtn submitBtnBig ClickShade']" ,6)
        sleep(5)'''
    testUrl = str(testWebUrl) + str(sheet["D" + str(i)].value).strip()
    sheet["D" + str(i)].value = testUrl
    webDriver.get(testUrl)
    sleep(3)
    html_source = webDriver.page_source
    if webDriver.current_url == str(testUrl):
        if textCheck in html_source:
            sheet["E" + str(i)].value = "有"
            sheet_date()
        else:
            sheet["E" + str(i)].value = "沒有"
            sheet_date()
    else:
        sheet["E" + str(i)].value = "無此url"
        sheet_date()
        
wb.save(str(testSiteID) + "_" + "JS檢查報告_chrome_H5_" + str(time.strftime("%y_%m_%d_%H_%M_%S") + ".xlsx"))
webDriver.quit()
