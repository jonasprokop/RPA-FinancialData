from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import PyPDF2
import re
import os
import time
import csv
# chrome driver setup
ser = Service('C:\\Users\\jonas\\PycharmProjects\\selen\\chromedriver.exe')
options = webdriver.ChromeOptions()
options.headless = False
options.add_argument("--window-size=1920,1200")
driver = webdriver.Chrome(service=ser, options=options)
# initiate download
driver.get("https://www.kurzy.cz/pdf/www.pse.cz/download-report/Issuers.dta/Emitenti2/CEZL012021.pdf")
driver.find_element(By.XPATH, '//*[@id="leftcolumn"]/div[2]/a[2]')
# wait for the file to download
downloadcheck = True
while downloadcheck == True:
    time.sleep(0.2)
    if os.path.exists('C:\\Users\\jonas\\Downloads\\CEZL012021.pdf') == True:
            downloadcheck = False
            driver.close()
# creating a pdf file object
pdfFileObj = open('C:\\Users\\jonas\\Downloads\\CEZL012021.pdf', 'rb')

# creating a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

# defining keywork
search_pattern1 = re.compile('EBITDA')

# define page range
pages=(i for i in range(99,100))

# extract text and do the search
for i in pages:

    PageObj = pdfReader.getPage(i)

    Text = PageObj.extractText()
    selectedIndex=1

    ResSearchVals = re.finditer (search_pattern1, Text)
    flag = 0
    for ResSearch in ResSearchVals:
        if(selectedIndex==flag):
            text=Text[ResSearch.span()[0]-20:ResSearch.span()[1]+20]
            if __name__ == '__main__':
                ls=re.split("\s{2,}|\\n",text)
        flag=flag+1
    number = ls.index ("EBITDA")
    target1 = ls[number+1]
    target2 = ls[number+2]
    print (target1 + ", " + target2)

# loads targets into csv output file

with open('C:\\Users\\jonas\\Downloads\\investicni_portfolio.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Spolecnost", "EBIT 1\6 2021", "EBIT 1\6 2020"])
    writer.writerow(["Cez", target1, target2])
