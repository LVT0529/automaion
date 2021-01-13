#selenium, openpyxl, pillow 모듈 설치 필요
#pip install selenium
#pip install openpyxl
#pip install Pillow

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time

startdate = input("(ex 2021-01-01 00:00) 시작 날짜를 입력 하세요: ")
enddate = input("(ex 2021-01-02 00:00) 종료 날짜를 입력 하세요: ")
A = [x for x in input("(여러개 장비 입력시 공백으로 구분하여 입력 ex 481 482) 확인할 장비 입력 바랍니다 : ").split()]
template = input("선택할 template 값을 입력 하세요: ")
search = input("(ex 검색 단어 없을시 그냥 엔터)검색할 단어를 입력 하세요: ")

driver = webdriver.Chrome()
driver.maximize_window()
driver.get("http://211.49.171.30/cacti/index.php")


username = "admin"
password = "7msanwk"

elem = driver.find_element_by_name('login_username')
elem.send_keys(username)
elem = driver.find_element_by_name("login_password")
elem.send_keys(password)
elem.send_keys(Keys.RETURN)

#Graph 이미지 클릭
driver.find_element_by_xpath('//*[@id="tabs"]/a[2]/img').click()

select = Select(driver.find_element_by_id("graph_template_id"))
select.select_by_value(template)

#시작 시간 설정
driver.find_element_by_id("date1").clear()
date = driver.find_element_by_id("date1")
date.send_keys(startdate)

#종료 시간 설정
driver.find_element_by_id("date2").clear()
date = driver.find_element_by_id("date2")
date.send_keys(enddate)
driver.find_element_by_name('button_refresh_x').click()

elem = driver.find_element_by_name("filter")
elem.send_keys(search)
elem.send_keys(Keys.RETURN)


count = 0
list = []
for host in A:
    for i in range(2,35):
        for j in range (1,4):
            try:
                select = Select(driver.find_element_by_id("host_id"))
                select.select_by_value(host)
                driver.find_element_by_xpath('//*[@id="main"]/table[2]/tbody/tr/td/table/tbody/tr[%s]/td[%s]/table/tbody/tr/td[2]/a[1]/img' %(i,j)).click()
                time.sleep(1)
                element1 = driver.find_element_by_class_name('graphimage')
                element_png = element1.screenshot_as_png
                list.append("%s-%s-%s.png" %(host,i,j))
                with open("%s-%s-%s.png" %(host,i,j),"wb") as file:
                    file.write(element_png)
                driver.back()
            except:
                break
                time.sleep(2)

#################################엑셀#################################
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Font
import openpyxl
import datetime


wb = openpyxl.Workbook()
ws = wb.active

ws = wb['Sheet']
ws.title = 'Cacti 캡쳐'


num = 1
B = ['A', 'J', 'S']
for i in list:
    for j in B:
        img = Image(i)
        ws.add_image(img, '%s%d' %(j,num))
    num += 12

nowdate = datetime.datetime.now()

wb.save(nowdate.strftime("%Y-%m-%d") + ' Cacti 이미지 캡쳐.xlsx')
