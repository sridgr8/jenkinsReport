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

testPlanName = sys.argv[1]
buildUrl = sys.argv[2]
executiondate=sys.argv[3]
buildUser=sys.argv[4]

today = date.today()
dayOfWeek = today.weekday()

print("Execution Started")
driver=webdriver.Chrome(executable_path="D:\\SELENIUM\\Chrome Driver\\chromedriver.exe")
driver.implicitly_wait(10)
driver.maximize_window()
driver.get(buildUrl+"timestamps/?elapsed=HH:mm:ss.S")

txtUN=driver.find_element_by_name("j_username")
txtUN.clear()
txtUN.send_keys("sridgr8")

txtPW=driver.find_element_by_name("j_password")
txtPW.clear()
txtPW.send_keys("purgatory")

txtPW.send_keys(Keys.RETURN)

txtTimeContent=str(driver.find_element_by_tag_name("pre").text)
txtTimeElapsed=txtTimeContent.splitlines()[-1]


print(txtTimeContent)
print("Total Time: "+txtTimeElapsed)

driver.get(buildUrl+"logText/progressiveText?start=0")

txtContent=str(driver.find_element_by_tag_name("pre").text)
txtContentLastLine=txtContent.splitlines()[-1]
txtBuildStatus=txtContentLastLine.split()[-1]
print("Build Status: "+txtBuildStatus)

filename1 = 'D:\\ECLIPSE PROJECTS\\TestRepoOne\\PythonExamples\\firstFile.xlsx'
wb = openpyxl.load_workbook(filename1)
ws = wb.active
rowNum=ws.max_row+1

now = datetime.today().strftime('%d-%m-%Y')
ws.cell(column=1, row=rowNum, value=testPlanName)
ws.cell(column=2, row=rowNum, value=txtTimeElapsed)
ws.cell(column=3, row = rowNum, value = now)
ws.cell(column=4, row= rowNum, value= txtBuildStatus)
ws.cell(column=5, row = rowNum, value= buildUser)
wb.save(filename1)

filename2 = 'D:\\ECLIPSE PROJECTS\\TestRepoOne\\PythonExamples\\secondFile.xlsx'
wb = openpyxl.load_workbook(filename2)
ws = wb.active
rowNum=ws.max_row+1
testPlanFound = 0
for x in range(2,rowNum):
    testPlanName1=ws.cell(row = x, column = 1)
    if (testPlanName1.value == testPlanName):
        testPlanFound = 1
        if(txtBuildStatus=="Passed"):
            passCount=ws.cell(row = x, column = 6)
            passCountVal=passCount.value
            ws.cell(column=6, row=x, value=int(passCountVal+1))
        elif(txtBuildStatus=="Failed"):
            failCount=ws.cell(row = x, column = 7)
            failCountVal=failCount.value
            ws.cell(column=7, row=x, value=int(failCountVal+1))
        else:
            otherCount=ws.cell(row = x, column = 8)
            otherCountVal=otherCount.value
            ws.cell(column=8, row=x, value=int(otherCountVal+1))
        wb.save(filename2)
if (testPlanFound == 0):
    ws.cell(column=1, row=rowNum, value=testPlanName)
    ws.cell(column=2, row=rowNum, value=txtTimeElapsed)
    ws.cell(column=3, row = rowNum, value = now)
    ws.cell(column=4, row= rowNum, value= txtBuildStatus)
    ws.cell(column=5, row = rowNum, value= builduser)
    if(txtBuildStatus=="Passed"):
        ws.cell(column=6, row=rowNum, value=1)
        ws.cell(column=7, row=rowNum, value=0)
        ws.cell(column=8, row=rowNum, value=0)
    elif(txtBuildStatus=="Failed"):        
        ws.cell(column=6, row=rowNum, value=0)
        ws.cell(column=7, row=rowNum, value=1)
        ws.cell(column=8, row=rowNum, value=0)
    else:
        ws.cell(column=6, row=rowNum, value=0)
        ws.cell(column=7, row=rowNum, value=0)
        ws.cell(column=8, row=rowNum, value=1)
    wb.save(filename2)
    
if (dayOfWeek == 3):
    filename3 = 'D:\\ECLIPSE PROJECTS\\TestRepoOne\\PythonExamples\\thirdFile.xlsx'
    wb = openpyxl.load_workbook(filename3)
    ws = wb.active
    rowNum=ws.max_row+1
    ws.cell(column=1, row=rowNum, value=testPlanName)
    ws.cell(column=2, row=rowNum, value=txtTimeElapsed)
    ws.cell(column=3, row = rowNum, value = now)
    ws.cell(column=4, row= rowNum, value= txtBuildStatus)
    ws.cell(column=5, row = rowNum, value= builduser)
    wb.save(filename3)
    filename4 = 'D:\\ECLIPSE PROJECTS\\TestRepoOne\\PythonExamples\\fourthFile.xlsx'
    wb = openpyxl.load_workbook(filename4)
    ws = wb.active
    rowNum=ws.max_row+1
    testPlanFound = 0
    for x in range(2,rowNum):
        testPlanName1=ws.cell(row = x, column = 1)
        if (testPlanName1.value == testPlanName):
            testPlanFound = 1
            if(txtBuildStatus=="Passed"):
                passCount=ws.cell(row = x, column = 6)
                passCountVal=passCount.value
                ws.cell(column=6, row=x, value=int(passCountVal+1))
            elif(txtBuildStatus=="Failed"):
                failCount=ws.cell(row = x, column = 7)
                failCountVal=failCount.value
                ws.cell(column=7, row=x, value=int(failCountVal+1))
            else:
                otherCount=ws.cell(row = x, column = 8)
                otherCountVal=otherCount.value
                ws.cell(column=8, row=x, value=int(otherCountVal+1))
            wb.save(filename4)
    if (testPlanFound == 0):
        ws.cell(column=1, row=rowNum, value=testPlanName)
        ws.cell(column=2, row=rowNum, value=txtTimeElapsed)
        ws.cell(column=3, row = rowNum, value = now)
        ws.cell(column=4, row= rowNum, value= txtBuildStatus)
        ws.cell(column=5, row = rowNum, value= builduser)
        if(txtBuildStatus=="Passed"):
            ws.cell(column=6, row=rowNum, value=1)
            ws.cell(column=7, row=rowNum, value=0)
            ws.cell(column=8, row=rowNum, value=0)
        elif(txtBuildStatus=="Failed"):        
            ws.cell(column=6, row=rowNum, value=0)
            ws.cell(column=7, row=rowNum, value=1)
            ws.cell(column=8, row=rowNum, value=0)
        else:
            ws.cell(column=6, row=rowNum, value=0)
            ws.cell(column=7, row=rowNum, value=0)
            ws.cell(column=8, row=rowNum, value=1)
        wb.save(filename4)    
    
print("Execution Completed")
