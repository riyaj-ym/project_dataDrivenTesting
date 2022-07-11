from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from xlUtils import xlUtils
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

ser_obj = Service('C:/Users/USER/Desktop/All Drivers/chromedriver')
driver = webdriver.Chrome(service=ser_obj)
driver.implicitly_wait(10)

driver.get('https://scripbox.com/plan/fd-calculator')

driver.maximize_window()

file = 'C:/Users/USER/Desktop/New folder/selenium/Selenium Project/calculator.xlsx'

sheetName = 'Sheet1'

rows = xlUtils.getRowCount(file, sheetName)
# print(rows)

for r in range(2, rows + 1):
    prAmt = xlUtils.readData(file, sheetName, r, 1)
    rateInt = xlUtils.readData(file, sheetName, r, 2)
    pr1 = xlUtils.readData(file, sheetName, r, 3)
    pr2 = xlUtils.readData(file, sheetName, r, 4)
    compPrd = xlUtils.readData(file, sheetName, r, 6)
    matExpected = xlUtils.readData(file, sheetName, r, 7)
    # print(pr2)
    # Passing Data to application
    a = driver.find_element(By.XPATH, '//*[@id="calcs"]/div/div/div[1]/div[3]/div[1]/div[1]/span/input')
    a.clear()
    a.send_keys(prAmt)
    b = driver.find_element(By.XPATH, '//*[@id="calcs"]/div/div/div[1]/div[3]/div[2]/div[1]/span/input')
    b.clear()
    b.send_keys(pr1)
    c = driver.find_element(By.XPATH, '//*[@id="calcs"]/div/div/div[1]/div[3]/div[3]/div[1]/span/input')
    c.clear()
    c.send_keys(rateInt)
    Prd = Select(driver.find_element(By.XPATH, '//*[@id="calcs"]/div/div/div[1]/div[3]/div[2]/div[1]/span/select'))
    Prd.select_by_visible_text(pr2)
    prd1 = Select(driver.find_element(By.XPATH, '//*[@id="calcs"]/div/div/div[1]/div[3]/div[4]/div/select'))
    prd1.select_by_visible_text(compPrd)
    matActual = driver.find_element(By.XPATH, '//*[@id="chart-div"]/div[1]/div[1]/p[2]').text

    matActual1 = matActual.replace('₹', '')
    matActual1 = matActual1.replace(',', '')
    matActual1 = matActual1.strip()

    print(matExpected)
    print(matActual1)

    if float(matExpected) == float(matActual1):
        print('Pass')
        xlUtils.writeData(file, sheetName, r, 9, 'Pass')
        xlUtils.fillGreenColor(file, sheetName, r, 9)
    else:
        print('Fail')
        xlUtils.writeData(file, sheetName, r, 9, 'Fail')
        xlUtils.fillRedColor(file, sheetName, r, 9)

time.sleep(10)
driver.close()
