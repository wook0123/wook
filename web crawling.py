from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from openpyxl import Workbook

url = 'https://finance.naver.com/sise/sise_rise.naver'

driver = webdriver.Chrome()
driver.implicitly_wait(3)
driver.get(url)
driver.find_element(By.ID, 'option2').click()
driver.find_element(By.ID, 'option8').click()
driver.find_element(By.ID, 'option14').click()
driver.find_element(By.ID, 'option20').click()
driver.find_element(By.ID, 'option4').click()
driver.find_element(By.ID, 'option5').click()
driver.find_element(By.ID, 'option17').click()
driver.find_element(By.ID, 'option24').click()
driver.find_element(By.CSS_SELECTOR, '#contentarea_left > div.box_type_m > form > div > div > div > a').click()

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
source = soup.select('#contentarea > div.box_type_l > table > tbody > tr')

write_wb = Workbook()
write_ws = write_wb.create_sheet('코스피')

datas = []
for n in source:
    data = []
    for j in n:
        trs = j.text.strip()
        data.append(trs)
    write_ws.append(data)


write_wb.save(r"C:\Users\USER\PycharmProjects\PythonProject4/연습3.xlsx")

driver.close()
