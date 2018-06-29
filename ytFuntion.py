from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from openpyxl import workbook ,load_workbook ,Workbook
import os ,time ,random ,ytFuntion

testdayFile = time.strftime("%y_%m_%d")
testdayTime  = time.strftime("%y_%m_%d_%H_%M_%S")

funtionError = []
funtionCountPng = 1

class accountSetting():
    def __init__(self ,username = "" ,password = "" ,safePassword = ""):
        self.username = str(username).strip()
        self.password = str(password).strip()
        self.safePassword = str(safePassword).strip()

class LocalStorage():
    def __init__(self ,webDriver):
        self.webDriver = webDriver    

    def __len__(self):
        return self.webDriver.execute_script("return window.localStorage.length;")

    def items(self) :
        return self.webDriver.execute_script( \
            "var ls = window.localStorage, items = {}; " \
            "for (var i = 0, k; i < ls.length; ++i) " \
            "  items[k = ls.key(i)] = ls.getItem(k); " \
            "return items; ")

    def keys(self) :
        return self.webDriver.execute_script( \
            "var ls = window.localStorage, keys = []; " \
            "for (var i = 0; i < ls.length; ++i) " \
            "  keys[i] = ls.key(i); " \
            "return keys; ")

    def get(self, key):
        return self.webDriver.execute_script("return window.localStorage.getItem(arguments[0]);", key)

    def set(self, key, value):
        self.webDriver.execute_script("window.localStorage.setItem(arguments[0], arguments[1]);", key, value)

    def has(self, key):
        return key in self.keys()

    def remove(self, key):
        self.webDriver.execute_script("window.localStorage.removeItem(arguments[0]);", key)

    def clear(self):
        self.webDriver.execute_script("window.localStorage.clear();")

    def __getitem__(self, key) :
        value = self.get(key)
        if value is None :
          raise KeyError(key)
        return value

    def __setitem__(self, key, value):
        self.set(key, value)

    def __contains__(self, key):
        return key in self.keys()

    def __iter__(self):
        return self.items().__iter__()

    def __repr__(self):
        return self.items().__str__()
    
class test_web(LocalStorage):
    def __init__(self ,webDriver):
        super().__init__(webDriver)

    def periodConfirm(self):
        try:
            self.webDriver.find_element_by_xpath("//span[.='确定']").click()
        except:
            pass

    def webItem(self):
        return self.webDriver.find_element_by_css_selector("ul[class='betFilter']").find_elements_by_tag_name('li') #取得item,固定寫法

    def webItemClick(self ,i):
        self.periodConfirm()
        if self.webItem()[i].text != "二同号单选":
            self.webItem()[i].click()

    def webPage(self):
        return self.webDriver.find_element_by_css_selector("ul[class='betNav fix']").find_elements_by_tag_name('li') #取得分頁,固定寫法

    def webPageClick(self ,i ,elementText = "" ,link_type = None):
        self.periodConfirm()
        self.webPage()[i].click()
        if i >= 5 and i < len(self.webPage()) and len(self.webPage()) > 6:
            self.elementClick(elementText ,link_type)

    def savePng(self ,save_Text = None ,drop_Down_count = "" ,donot_Save = ""):
        if save_Text == None or str(donot_Save) != "":
            return
        global funtionError ,funtionCountPng  #全域變數被當成區域變數的解法
        web_Height = self.webDriver.execute_script("return document.body.scrollHeight")
        webPosition_y ,i ,drop_Down = 0 ,1 ,1
        while i <= drop_Down:
            try:
                self.webDriver.execute_script("window.scroll(0, "+ str(webPosition_y) +");")
                sleep(1)
                self.periodConfirm()
                self.webDriver.save_screenshot(testdayFile + "/" + str(testdayTime) + "_" + str(funtionCountPng) + "_" + str(save_Text) + ".png")
                funtionCountPng += 1
                webPosition_y += 600
                i += 1
                if drop_Down_count != "":
                    drop_Down = int(drop_Down_count)
                else:
                    drop_Down = (self.webDriver.execute_script("return document.body.scrollHeight") / 600) + 1
            except:
                funtionError.append(save_Text + str(funtionCountPng) + "_NG")
                return funtionError

    def elementClick(self ,elementText = "" ,link_type = None ,delayTime = 0):
        global funtionError #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            delayTime = int(str(delayTime).strip())
            if link_type == 1:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.ID,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_id(elementText).click()
            elif link_type == 2:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_class_name(elementText).click()
            elif link_type == 3:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_link_text(elementText).click()
            elif link_type == 4:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_partial_link_text(elementText).click()
            elif link_type == 5:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_name(elementText).click()
            elif link_type == 6:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_css_selector(elementText).click()
            elif link_type == 7:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.TAG_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_tag_name(elementText).click()
            elif link_type == 8:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.XPATH,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_xpath(elementText).click()
            else:
                return funtionError.append(elementText + "_" + str(link_type) + "_ClickNG")
        except:
            funtionError.append(elementText + "_" + str(link_type) + "_NG")
            return funtionError

    def elementsClickOne(self ,elementText = "",link_type = None ,elements_num = 0 ,delayTime = 0):
        global funtionError #全域變數被當成區域變數的解法
        self.periodConfirm()
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            elements_num = int(str(elements_num).strip()) -1
            if link_type == 1:
                self.webDriver.find_elements_by_id(elementText)[elements_num].click()
            elif link_type == 2:
                self.webDriver.find_elements_by_class_name(elementText)[elements_num].click()
            elif link_type == 3:
                self.webDriver.find_elements_by_link_text(elementText)[elements_num].click()
            elif link_type == 4:
                self.webDriver.find_elements_by_partial_link_text(elementText)[elements_num].click()
            elif link_type == 5:
                self.webDriver.find_elements_by_name(elementText)[elements_num].click()
            elif link_type == 6:
                self.webDriver.find_elements_by_css_selector(elementText)[elements_num].click()
            elif link_type == 7:
                self.webDriver.find_elements_by_tag_name(elementText)[elements_num].click()
            elif link_type == 8:
                self.webDriver.find_elements_by_xpath(elementText)[elements_num].click()
            else:
                return funtionError.append(elementText + "_" + str(link_type) + "_ClickNG")
        except:
            funtionError.append(elementText + "_" + str(link_type) + "_NG")
            return funtionError

    def elementsClickAll(self ,elementText = "",link_type = None ,elements_num = 0 ,delayTime = 0):
        global funtionError #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            elements_num = int(str(elements_num).strip())
            sleep(int(delayTime))
            if link_type == 1:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_id(elementText)[i].click()
            elif link_type == 2:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_class_name(elementText)[i].click()
            elif link_type == 3:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_link_text(elementText)[i].click()
            elif link_type == 4:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_partial_link_text(elementText)[i].click()
            elif link_type == 5:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_name(elementText)[i].click()
            elif link_type == 6:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_css_selector(elementText)[i].click()
            elif link_type == 7:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_tag_name(elementText)[i].click()
            elif link_type == 8:
                for i in range(elements_num):
                    self.periodConfirm()
                    self.webDriver.find_elements_by_xpath(elementText)[i].click()
            else:
                return funtionError.append(elementText + "_" + str(link_type) + "_ClickNG")
        except:
            funtionError.append(elementText + "_" + str(link_type) + "_NG")
            return funtionError

    def elements(self ,elementText = "" ,link_type = None ,delayTime = 0):
        global funtionError #全域變數被當成區域變數的解法
        self.periodConfirm()
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            sleep(int(delayTime))
            if link_type == 1:
                return self.webDriver.find_elements_by_id(elementText)
            elif link_type == 2:
                return self.webDriver.find_elements_by_class_name(elementText)
            elif link_type == 3:
                return self.webDriver.find_elements_by_link_text(elementText)
            elif link_type == 4:
                return self.webDriver.find_elements_by_partial_link_text(elementText)
            elif link_type == 5:
                return self.webDriver.find_elements_by_name(elementText)
            elif link_type == 6:
                return self.webDriver.find_elements_by_css_selector(elementText)
            elif link_type == 7:
                return self.webDriver.find_elements_by_tag_name(elementText)
            elif link_type == 8:
                return self.webDriver.find_elements_by_xpath(elementText)
            else:
                funtionError.append(elementText + "_" + str(link_type) + "_get_NG")
                return funtionError
        except:
            funtionError.append(elementText + "_" + str(link_type) + "_get_NG")
            return funtionError

    def element(self ,elementText = "",link_type = None):
        global funtionError #全域變數被當成區域變數的解法
        self.periodConfirm()
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            if link_type == 1:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.ID,elementText)))
                return self.webDriver.find_element_by_id(elementText)
            elif link_type == 2:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME,elementText)))
                return self.webDriver.find_element_by_class_name(elementText)
            elif link_type == 3:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,elementText)))
                return self.webDriver.find_element_by_link_text(elementText)
            elif link_type == 4:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                return self.webDriver.find_element_by_partial_link_text(elementText)
            elif link_type == 5:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.NAME,elementText)))
                return self.webDriver.find_element_by_name(elementText)
            elif link_type == 6:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,elementText)))
                return self.webDriver.find_element_by_css_selector(elementText)
            elif link_type == 7:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.TAG_NAME,elementText)))
                return self.webDriver.find_element_by_tag_name(elementText)
            elif link_type == 8:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.XPATH,elementText)))
                return self.webDriver.find_element_by_xpath(elementText)
            else:
                funtionError.append(elementText + "_" + str(link_type) + "_ClickNG")
                return funtionError
        except:
            funtionError.append(elementText + "_" + str(link_type) + "_NG")
            return funtionError

    def elementSendKeys(self ,elementText = "" ,link_type = None ,delayTime = 0 ,text = ""):
        global funtionError #全域變數被當成區域變數的解法
        try:
            link_type = int(str(link_type).strip())
            elementText = str(elementText).strip()
            delayTime = int(str(delayTime).strip())
            text = str(text).strip()
            if link_type == 1:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.ID,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_id(elementText).send_keys(text)
            elif link_type == 2:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_class_name(elementText).send_keys(text)
            elif link_type == 3:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_link_text(elementText).send_keys(text)
            elif link_type == 4:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_partial_link_text(elementText).send_keys(text)
            elif link_type == 5:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_name(elementText).send_keys(text)
            elif link_type == 6:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_css_selector(elementText).send_keys(text)
            elif link_type == 7:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.TAG_NAME,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_tag_name(elementText).send_keys(text)
            elif link_type == 8:
                WebDriverWait(self.webDriver, 10).until(EC.visibility_of_element_located((By.XPATH,elementText)))
                if delayTime != 0:
                    sleep(delayTime)
                self.periodConfirm()
                self.webDriver.find_element_by_xpath(elementText).send_keys(text)
            else:
                return funtionError.append(elementText + "_" + str(link_type) + "_Send_key_NG")
        except:
            funtionError.append(elementText + "_" + str(link_type) + "_NG")
            return funtionError

    def rebate(self ,elementText = "" ,link_type = None ,element2_Text = "" ,link2_type = None ,rebate_Number = ""):
        link_type = int(str(link_type).strip())
        elementText = str(elementText).strip()
        link2_type = int(str(link2_type).strip())
        element2_Text = str(element2_Text).strip()
        rebate_Number = str(self.element(elementText ,link_type).text)
        self.elementSendKeys(element2_Text ,link2_type ,text = rebate_Number[0:-2])
        return rebate_Number[0:-2]
    
    def periodDetail(self):
        periodDetail = []
        while(True):
            WebDriverWait(self.webDriver, 10).until(EC.visibility_of_all_elements_located((By.TAG_NAME,"td")))
            for i in range(len(self.elements("td" ,7))):
                try:
                    periodDetail.append(self.elements("td" ,7)[i].text)
                except:
                    pass
            try:
                self.webDriver.find_element_by_xpath("//a[.='下一页']").click()
            except:
                break
        return periodDetail

    def sheetDetail(self):
        periodDetail = self.periodDetail()
        for i in range(int(len(periodDetail)/8)):
            sheetDetail["B"+str(len(sheetDetail["B"]) + 1)].value = periodDetail[8*i]
            sheetDetail["C"+str(len(sheetDetail["B"]))].value = periodDetail[1 + 8*i]
            sheetDetail["D"+str(len(sheetDetail["B"]))].value = periodDetail[2 + 8*i]
            sheetDetail["E"+str(len(sheetDetail["B"]))].value = periodDetail[3 + 8*i]
            sheetDetail["F"+str(len(sheetDetail["B"]))].value = periodDetail[4 + 8*i]
            sheetDetail["G"+str(len(sheetDetail["B"]))].value = periodDetail[5 + 8*i]
            sheetDetail["H"+str(len(sheetDetail["B"]))].value = periodDetail[6 + 8*i]
            sheetDetail["I"+str(len(sheetDetail["B"]))].value = periodDetail[7 + 8*i]

    def speed_3_t_r(self ,elementText = "" ,link_type = None ,max_Td = "0" ,max_Money = "0"):
        money = ["金額"]
        link_type = str(link_type).strip()
        elementText = str(elementText).strip()
        money_box = self.elements(elementText ,link_type)
        max_Td = str(max_Td).strip()
        max_Money = str(max_Money).strip()
        if int(max_Td) == 0:
            money_box = money_box[0:-1]
        else:
            money_box = money_box[0:int(max_Td)]
        for i in range(1 ,len(money_box)):
            self.periodConfirm()
            money_box[i].clear()
            if int(max_Money) == 0:
                money.append(random.randint(0 ,99))
            else:
                money.append(int(max_Money))
            money_box[i].send_keys(str(money[i]))
        return money

    def speed_3_r(self ,elementText = "" ,link_type = None ,max_Td = "0" ,max_Money = "0"):
        money = ["投注"]
        link_type = str(link_type).strip()
        elementText = str(elementText).strip()
        money_box = self.elements(elementText ,link_type)
        max_Td = str(max_Td).strip()
        max_Money = str(max_Money).strip()
        if int(max_Td) == 0:
            money_box = money_box
        else:
            money_box = money_box[0:int(max_Td)]
        for i in range(len(money_box)):
            self.periodConfirm()
            money_box[i].clear()
            if int(max_Money) == 0:
                money.append(self.elements("order_type" ,2)[i].text)
                money.append(self.elements("order_zhushu" ,2)[i].text)
                money.append(random.randint(0 ,99))
            else:
                money.append(self.elements("order_type" ,2)[i].text)
                money.append(self.elements("order_zhushu" ,2)[i].text)
                money.append(int(max_Money))
            money_box[i].send_keys(str(money[3 + 3*i]))                                  
        return money

class sheet_work():
    def __init__(self ,sheet_work):
        self.sheet_work = sheet_work

    def sheet_value(self ,col = "" ,colNumber = "" ,value = ""):
        col = str(col).strip()
        colNumber = str(colNumber).strip()
        value = str(value).strip()
        self.sheet_work[col + str(len(self.sheet_work[colNumber]))].value = value
