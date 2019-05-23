import os
from selenium.webdriver.common.keys import Keys
from selenium import webdriver 
import time 
from openpyxl import load_workbook
wb = load_workbook('ID.xlsx')
sheet = wb.active 

#期交所登入
def CustomerID(CNAM,CID):
    path= CNAM + "\\"
    if not os.path.isdir(path):
    	os.mkdir(path)         
    
    driver = webdriver.Chrome()
    driver.set_window_size(1200, 700)
    driver.implicitly_wait(15)    
    driver.get("https://report.taifex.com.tw/FMS/login.html")
    
    User = driver.find_element_by_id("username")
    User.clear()
    User.send_keys("*****")
    
    ID = driver.find_element_by_id("subId")
    ID.clear()
    ID.send_keys("**")
    
    Passwd = driver.find_element_by_id("j_password")
    Passwd.clear()
    Passwd.send_keys("*******")                 

    #--------------------------------期交所違約-------------------------------------
    driver.switch_to.parent_frame()
    time.sleep(0.5)
    driver.find_element_by_id("j_password").send_keys(Keys.ENTER)
    driver.find_element_by_xpath('//*[@id="ext-gen42"]').click()
    driver.find_element_by_xpath('//*[@id="ext-gen48"]').click()
    driver.find_element_by_xpath('//*[@id="ext-gen49"]/a[2]').click()
    driver.switch_to.frame(driver.find_element_by_xpath('//*[@id="main"]'))
    time.sleep(0.5)
    driver.find_element_by_xpath('//*[@value=1]').click()
    driver.find_element_by_xpath('//*[@value="送出"]').click()
    time.sleep(0.5)
    
    ID = driver.find_element_by_name('someid')
    ID.clear()
    ID.send_keys(CID)
    
    driver.find_element_by_xpath('//*[@value="送出"]').click()
    time.sleep(0.5)

    driver.save_screenshot(path+CID+"_違約案件"+".png")

    #----------------------------------市場開戶數-----------------------------------
    driver.switch_to.parent_frame()
    
    time.sleep(0.5)
    driver.find_element_by_xpath('//*[@id="ext-gen54"]').click()
    driver.find_element_by_xpath('//*[@id="ext-gen55"]').click()
    driver.switch_to.frame(driver.find_element_by_xpath('//*[@id="main"]'))
    
    ID = driver.find_element_by_name('W_id1')
    ID.clear()
    ID.send_keys(CID)
    
    driver.find_element_by_xpath('//*[@value="送出資料"]').click()
    time.sleep(0.5)
 
    driver.save_screenshot(path+CID+"_市場開戶數"+".png")

    #----------------------------------家事事件-----------------------------------
    driver.get("http://domestic.judicial.gov.tw/abbs/wkw/WHD9HN01.jsp")
    
    ID = driver.find_element_by_name('idno')
    ID.clear()
    ID.send_keys(CID) 
    driver.find_element_by_xpath('//*[@value="　查　詢　"]').click()
    time.sleep(0.5)
    driver.save_screenshot(path+CID+"_家事事件"+".png")
    driver.close()
    return
     
for i in range(2,len(sheet[('A')])+1):
    i = str(i)
    CustomerID(sheet[('A'+i)].value,sheet[('B'+i)].value)
    
