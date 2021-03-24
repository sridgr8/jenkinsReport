'''
Created on March 23, 2021

@author: Srinivasulu
'''

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
import openpyxl 
import sys
import string
from datetime import date
import os

# testPlanName = sys.argv[1]
# buildUrl = sys.argv[2]
# executiondate=sys.argv[3]
# buildUser=sys.argv[4]

# today = date.today()
# dayOfWeek = today.weekday()

testPlanName="DAM Custom Execution"
buildDuration="30 seconds"

print("Execution Started")
# driver=webdriver.Chrome(executable_path="C:\\wamp64\\www\\jenkinsReport\\webDriver\\chromedriver.exe")
# driver.implicitly_wait(10)
# driver.maximize_window()
# driver.get(buildUrl+"timestamps/?elapsed=HH:mm:ss.S")

# filename1 = 'C:\\wamp64\\www\\jenkinsReport\\excelFiles\\testExcel.xlsx'
filename1 = str(os.getcwd())+'\\excelFiles\\testExcel.xlsx'
wb = openpyxl.load_workbook(filename1)
ws = wb.active
rowNum=ws.max_row+1

ws.cell(column=1, row=rowNum, value=testPlanName)
ws.cell(column=2, row=rowNum, value=buildDuration)
wb.save(filename1)
print("Execution Completed")
