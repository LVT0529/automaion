#selenium, openpyxl, pillow 모듈 설치 필요
#pip install selenium
#pip install openpyxl
#pip install Pillow

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import os

startdate = input("(ex 2021-01-01 00:00) 시작 날짜를 입력 하세요: ")
enddate = input("(ex 2021-01-02 00:00) 종료 날짜를 입력 하세요: ")
#template = input("선택할 template 값을 입력 하세요: ")
#A = [x for x in input("(여러개 장비 입력시 공백으로 구분하여 입력 ex 481 482) 확인할 장비 입력 바랍니다 : ").split()]
#search = input("(ex 검색 단어 없을시 그냥 엔터)검색할 단어를 입력 하세요: ")

# Headless Web 설정
'''
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument("disable-gpu")
options.add_argument('window-size=1920x1080')
options.add_argument("lang=ko_KR")
driver = webdriver.Chrome(options=options)
'''


driver = webdriver.Chrome()
driver.maximize_window()


elem = driver.find_element_by_name('login_username')
elem.send_keys(username)
elem = driver.find_element_by_name("login_password")
elem.send_keys(password)
elem.send_keys(Keys.RETURN)

#Graph 이미지 클릭
driver.find_element_by_xpath('//*[@id="tabs"]/a[2]/img').click()

#Total Template 선택
select = Select(driver.find_element_by_id("graph_template_id"))
select.select_by_value('70')

#시작 시간 설정
driver.find_element_by_id("date1").clear()
date = driver.find_element_by_id("date1")
date.send_keys(startdate)

#종료 시간 설정
driver.find_element_by_id("date2").clear()
date = driver.find_element_by_id("date2")
date.send_keys(enddate)
driver.find_element_by_name('button_refresh_x').click()


#검색
'''
elem = driver.find_element_by_name("filter")
elem.send_keys(search)
elem.send_keys(Keys.RETURN)
'''

now_time = time.strftime('%y-%m-%d')
print (now_time)
## Graph_Capture 폴더 생성
if not os.path.isdir("Graph_Capture " + now_time):
   os.makedirs("Graph_Capture " + now_time)
   print("Make directory 'Graph_Capture'")
os.chdir(os.getcwd() + "\Graph_Capture " + now_time)


graph_list = ['graph_27001', 'graph_14402', 'graph_14361', 'graph_26661', 'graph_23318',
              'graph_29243', 'graph_14907', 'graph_14924', 'graph_1008', 'graph_1485',
              'graph_11347', 'graph_11396', 'graph_27836', 'graph_27849']

list_before = []
count = 0

for tem in range (1,3):
    for num in range(0,len(graph_list)):
        for i in range(2,35):
            for j in range (1,4):
                try:
                    #select = Select(driver.find_element_by_id("host_id"))
                    #select.select_by_value(host)
                    #path = '//*[@id="main"]/table[2]/tbody/tr/td/table/tbody/tr[' + str(i) + ']/td['+ str(j) +']/table/tbody/tr/td[2]/a[1]/img'
                    #driver.find_element_by_xpath(path).click()
                    path = '//*[@id="main"]/table[2]/tbody/tr/td/table/tbody/tr[' + str(i) + ']/td['+ str(j) +']/table/tbody/tr/'
                    graph_id = driver.find_element_by_xpath(path + 'td[1]/div/a/img').get_attribute("id")

                    if(graph_id in graph_list[num]):
                        driver.find_element_by_xpath(path + 'td[2]/a[1]/img').click()
                        element1 = driver.find_element_by_class_name('graphimage')
                        element_png = element1.screenshot_as_png

                        graph_name = element1.get_attribute("alt")
                        graph_name = graph_name.split("[")

                        list_before.append("%s.png" %(graph_name[0]))
                        with open("%s.png" %(graph_name[0]),"wb") as file:
                            file.write(element_png)
                        driver.back()
                except:
                    break
                    time.sleep(1)
    # Global Connection Template 선택
    select = Select(driver.find_element_by_id("graph_template_id"))
    select.select_by_value('38')


driver.quit()

###########################
def capture(date1, date2):













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
for i in list_before:
    img = Image(i)
    ws.add_image(img, 'A%d' %(num))
    num += 12
'''
B = ['J']
for i in list_before:
    for j in B:
        img = Image(i)
        ws.add_image(img, '%s%d' %(j,num))
    num += 12
'''

nowdate = datetime.datetime.now()

wb.save(nowdate.strftime("%Y-%m-%d") + ' Cacti 이미지 캡쳐.xlsx')
