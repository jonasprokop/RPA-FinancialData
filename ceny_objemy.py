from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
# chrome driver setup
ser = Service('C:\\Users\\jonas\\PycharmProjects\\selen\\chromedriver.exe')
options = webdriver.ChromeOptions()
options.headless = False
options.add_argument("--window-size=1920,1200")
driver = webdriver.Chrome(service=ser, options=options)
# opens the page in browser
driver.get('https://www.kurzy.cz/akcie-cz/akcie/cez-183/')
# locates the element by XPATH
element1 = driver.find_element(By.XPATH, '//*[@id="cena_trhy"]/div[2]/table/tbody/tr[2]/td[2]/b')
# gets text from element
cena_akcie = element1.text
print(cena_akcie)

