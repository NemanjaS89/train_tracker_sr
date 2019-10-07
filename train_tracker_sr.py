from selenium import webdriver
from bs4 import BeautifulSoup
import xlrd
import xlsxwriter
import time

#napravi novi fajl i spremi ga za pisanje
workbook = xlsxwriter.Workbook('trains_sr_generated.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0

#otvori ulazni fajl za čitanje
loc = 'trains_sr.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

#otvaranje browsera
browser = webdriver.Chrome('resources/chromedriver.exe')
browser.get('http://portal.srbcargo.rs/kargoportal/Form1')

#petlja
for i in range (sheet.nrows):
    text_field = browser.find_element_by_name('TextBox1')
    text_field.send_keys(str(sheet.cell_value(i, 0)))
    
    submit_button = browser.find_element_by_name('btnInsert')
    submit_button.click()
    
    soup = BeautifulSoup(browser.page_source, 'lxml')
    status = soup.find('span', {'id': 'lblStatusVagona'})
    location = soup.find('span', {'id': 'lbl1'})
    time = soup.find('span', {'id': 'lbl2'})
    
    worksheet.write(row, col, str(sheet.cell_value(i, 0)))
    worksheet.write(row, col + 1, status.text)
    worksheet.write(row, col + 2, location.text)
    worksheet.write(row, col + 3, time.text)
    
    browser.find_element_by_name('TextBox1').clear()
    
    row += 1
    
workbook.close()