from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
# chrome driver setup
ser = Service('C:\\Users\\jonas\\PycharmProjects\\RPA_financni_informace\\chromedriver.exe')
options = webdriver.ChromeOptions()
options.headless = False
options.add_argument("--window-size=1920,1200")
driver = webdriver.Chrome(service=ser, options=options)
# opens the page in browser
driver.get('https://www.patria.cz/akcie/online/svet.html')
driver.maximize_window()
# find list of companies by segment
time.sleep(3)
driver.find_element(By.ID, 'onetrust-accept-btn-handler').click()
time.sleep(5)
driver.find_element(By.ID, 'ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_Sector_Input').click()
time.sleep(3)
driver.find_element(By.XPATH, '//*['
                              '@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_Sector_DropDown"]/div/ul/li[5]').click()
driver.find_element(By.ID, 'ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_PatriaQuery').click()
time.sleep(5)
# scrape the table of companies with basic metrics
tbl = driver.find_element(By.XPATH, '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_ctl13"]').get_attribute("outerHTML")
dfs = pd.read_html(tbl)
df = pd.concat(dfs)
dict = {'Akce': 'Action',
        'Název': 'Name',
        'Datum a čas': 'Datetime',
        'Change': 'Change',
        'Unnamed: 3': 'Smazat',
        'Nejlepšínákup': 'BestBuy',
        'Nejlepšíprodej': 'Bestsell',
        'Posledníobchod': 'LastTrade',
        'Změna(%)': 'Change',
        'Objem(ks)': 'Volume',
        'Měna': 'Currency',
        'Trh': 'Market',
        'Závěrečný kurz': 'CloseValue',
        }
df.rename(columns=dict,
          inplace=True)
# sorting prior to selection
df_sorted_volume = df.sort_values('Volume')
df_sorted_volume_100 = df_sorted_volume.head(100)
df_sorted_change = df_sorted_volume_100.sort_values('Change')
# selection
df_sorted_change_25 = df_sorted_change.head(25)
print (df_sorted_change_25)
list_of_firms = df_sorted_change_25['Name'].tolist()
# lists for metrics scraping
profit_on_share_list = []
EBITDA_list = []
enterprise_value_list = []
cash_flow_on_share_list = []
# loops through list of companies and loads selected metrics from webpage
for firm in list_of_firms:
    search = driver.find_element(By.ID, 'searchBarCondition')
    search.send_keys(firm)
    driver.find_element(By.ID, 'ctl00_ctl00_ctl00_MC_searchBarButton').click()
    driver.find_element(By.XPATH, '//*[@id="ctl00_ctl00_ctl00_MC_Content_OneColumnContent_SearchResultEquity"]/div[3]/div/div/table/tbody/tr[2]/td[1]/a').click()
    driver.find_element(By.XPATH, '//*[@id="rightColumn"]/div/div[3]/ul/li[5]/span/a').click()
    time.sleep(3)
    profit_on_share = driver.find_element(By.XPATH, '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[1]/td[2]').text
    EBITDA = driver.find_element(By.XPATH, '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[3]/td[2]').text
    enterprise_value = driver.find_element(By.XPATH, '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[6]/td[5]').text
    cash_flow_on_share = driver.find_element(By.XPATH, '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[7]/td[5]').text
    profit_on_share_list.append(profit_on_share)
    EBITDA_list.append(EBITDA)
    enterprise_value_list.append(enterprise_value)
    cash_flow_on_share_list.append(cash_flow_on_share)
# values of scraped metric were preloaded into set of lists
# lists are loaded into pandas dataframe
prehled_firem = pd.DataFrame()
prehled_firem['Jméno']= list_of_firms
prehled_firem['Profit na akcii']= profit_on_share_list
prehled_firem['EBITDA']= EBITDA_list
prehled_firem['Cena Enterprise']= cash_flow_on_share_list
prehled_firem['Cashflow na akcii']= profit_on_share_list

driver.close()
print (prehled_firem)
prehled_firem.to_excel('C:\\Users\\jonas\\PycharmProjects\\RPA_financni_informace\\Seznam_firem.xls', sheet_name='Prehled firem')
