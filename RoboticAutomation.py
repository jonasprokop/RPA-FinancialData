from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd
import numpy as np
import warnings

class RoboticProcessAutomation:
    # Searches for list of candidates for investment, and then scrapes financial data about them
    def __init__(self, n, path):
        self.n = n
        self.path = path
        self.open_driver()
        self.get_table()
        self.selection(self.n)
        self.company_loop()
        self.pandas_to_excel(self.path)
        self.driver_shutdown()

    # chrome driver setup variables
    ser = Service('C:\\Users\\jonas\\PycharmProjects\\RPA_financni_informace\\chromedriver.exe')
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=ser, options=options)

    # lists and dataframes for metrics scraping
    names_list = []
    changes_list = []
    volumes_list = []
    last_trade_list = []
    price_list = []
    change_list = []
    equity_list = []
    market_capitalization_list = []
    total_turnover_list = []
    EBITDA_list = []
    pure_gain_for_shareholders_list = []
    gain_on_share_list = []
    profit_on_share_list = []
    BV_list = []
    cash_on_share_list = []
    gross_margin_list = []
    ROE_list = []
    P_E_list = []
    pandas_list = []
    price_sales_on_share_list = []
    P_BV_list = []
    enterprise_value_list = []
    cash_flow_on_share_list = []
    dividends_on_share_list = []
    list_of_companies = []
    company_comparison = pd.DataFrame()
    overview_of_companies = pd.DataFrame()

    def open_driver(self):
        # open driver with the target webpage and accept conditions
        warnings.simplefilter(action='ignore', category=FutureWarning)
        self.options.headless = False
        self.options.add_argument('--window-size=1920,1200')
        self.driver.get('https://www.patria.cz/akcie/online/svet.html')
        self.driver.maximize_window()
        time.sleep(3)
        self.driver.find_element(By.ID, 'onetrust-accept-btn-handler').click()

    def get_table(self):
        # lists companies in selected industry
        # scrapes companies names and basic metrics into table
        time.sleep(5)
        self.driver.find_element(By.ID,
                                 'ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_Sector_Input').click()
        time.sleep(3)

        self.driver.find_element(By.XPATH, '//*['
                                           '@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_Sector_DropDown"]/div/ul/li[5]').click()
        self.driver.find_element(By.ID,
                                 'ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_PatriaQuery').click()
        time.sleep(5)
        rows = len(self.driver.find_elements(By.XPATH,
                                             '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_ctl13"]/tbody/tr'))
        for company in range(2, rows):
            names = self.driver.find_element(By.XPATH,
                                             f'//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_ctl13"]/tbody/tr[{company}]/td[2]/a').text
            changes = self.driver.find_element(By.XPATH,
                                               f'//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_ctl13"]/tbody/tr[{company}]/td[8]').text
            volumes = self.driver.find_element(By.XPATH,
                                               f'//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_AdvancedSearchBig_ctl13"]/tbody/tr[{company}]/td[9]').text
            self.names_list.append(names)
            self.changes_list.append(changes)
            self.volumes_list.append(volumes)
        self.company_comparison['Name'] = self.names_list
        self.company_comparison['Change'] = self.changes_list
        self.company_comparison['Volume'] = self.volumes_list

    def selection(self, n):
        # filters and orders the table of companies and then selects list of n of them for further research
        self.company_comparison.replace('-', np.nan, inplace=True)
        self.company_comparison.replace('', np.nan, inplace=True)
        self.company_comparison['Change'] = self.company_comparison['Change'].str.replace(',', '.')
        self.company_comparison['Change'] = self.company_comparison['Change'].str.replace(' ', '')
        self.company_comparison['Volume'] = self.company_comparison['Volume'].str.replace(',', '.')
        self.company_comparison['Volume'] = self.company_comparison['Volume'].str.replace(' ', '')
        self.company_comparison.dropna(subset=['Name'], inplace=True)
        self.company_comparison.dropna(subset=['Volume'], inplace=True)
        self.company_comparison.dropna(subset=['Change'], inplace=True)
        pd.to_numeric(self.company_comparison['Volume'], errors='coerce', downcast='float')
        pd.to_numeric(self.company_comparison['Change'], errors='coerce', downcast='float')
        company_comparison_float = self.company_comparison.astype({'Volume': float, 'Change': float})
        company_comparison_float.dropna(subset=['Change'], inplace=True)
        company_comparison_float.dropna(subset=['Volume'], inplace=True)
        index_names_plus = company_comparison_float.index[company_comparison_float['Change'] >= 75.0]
        company_comparison_float.drop(index_names_plus, inplace=True)
        index_names_minus = company_comparison_float.index[company_comparison_float['Change'] <= 0.0]
        company_comparison_float.drop(index_names_minus, inplace=True)
        company_comparison_float.sort_values(by=['Change'], ascending=False, ignore_index=True, inplace=True)
        company_comparison_selected = company_comparison_float.head(n + 1)
        self.list_of_companies = company_comparison_selected['Name'].tolist()

    def company_loop(self):
        # loops through list of companies
        # for each acceses webpage, scrapes and loads target data into dataframe
        for company_ in self.list_of_companies:
            search = self.driver.find_element(By.ID, 'searchBarCondition')
            search.send_keys(company_)
            self.driver.find_element(By.ID, 'ctl00_ctl00_ctl00_MC_searchBarButton').click()
            self.driver.find_element(By.XPATH,
                                     '//*[@id="ctl00_ctl00_ctl00_MC_Content_OneColumnContent_SearchResultEquity"]/div[3]/div/div/table/tbody/tr[2]/td[1]/a').click()
            self.driver.find_element(By.XPATH, '//*[@id="rightColumn"]/div/div[3]/ul/li[5]/span/a').click()
            time.sleep(5)
            overview_scrape_1 = []
            overview_scrape_2 = []
            overview_scrape_1 = self.driver.find_elements(By.CLASS_NAME, 'EquityHeaderCellStrong')
            overview_scrape_2 = self.driver.find_elements(By.CLASS_NAME, 'EquityHeaderCell')
            last_trade_price = []
            change_equity = []
            for i in overview_scrape_1:
                last_trade_price.append(i.text)
            for i in overview_scrape_2:
                change_equity.append(i.text)
            market_capitalization = self.driver.find_element(By.XPATH,
                                                             '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[1]/td[2]').text
            total_turnover = self.driver.find_element(By.XPATH,
                                                      '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[2]/td[2]').text
            EBITDA = self.driver.find_element(By.XPATH,
                                              '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[3]/td[2]').text
            pure_gain_for_shareholders = self.driver.find_element(By.XPATH,
                                                                  '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[4]/td[2]').text
            gain_on_share = self.driver.find_element(By.XPATH,
                                                     '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[4]/td[2]').text
            profit_on_share = self.driver.find_element(By.XPATH,
                                                       '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[1]/td[2]').text
            BV = self.driver.find_element(By.XPATH,
                                          '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[7]/td[2]').text
            cash_on_share = self.driver.find_element(By.XPATH,
                                                     '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[8]/td[2]').text
            gross_margin = self.driver.find_element(By.XPATH,
                                                    '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[1]/td[5]').text
            ROE = self.driver.find_element(By.XPATH,
                                           '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[2]/td[5]').text
            price_sales_on_share = self.driver.find_element(By.XPATH,
                                                            '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[3]/td[5]').text
            P_E = self.driver.find_element(By.XPATH,
                                           '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[4]/td[5]').text
            P_BV = self.driver.find_element(By.XPATH,
                                            '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[5]/td[5]').text
            enterprise_value = self.driver.find_element(By.XPATH,
                                                        '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[6]/td[5]').text
            cash_flow_on_share = self.driver.find_element(By.XPATH,
                                                          '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[7]/td[5]').text
            dividends_on_share = self.driver.find_element(By.XPATH,
                                                          '//*[@id="ctl00_ctl00_ctl00_MC_Content_rightColumnPlaceHolder_DetailCorporateEconomyResults"]/div[3]/div[1]/table[3]/tbody/tr[8]/td[5]').text

            self.last_trade_list.append(last_trade_price[0])
            self.price_list.append(last_trade_price[1])
            self.change_list.append(change_equity[0])
            self.equity_list.append(change_equity[1])
            self.market_capitalization_list.append(market_capitalization)
            self.total_turnover_list.append(total_turnover)
            self.EBITDA_list.append(EBITDA)
            self.pure_gain_for_shareholders_list.append(pure_gain_for_shareholders)
            self.gain_on_share_list.append(gain_on_share)
            self.profit_on_share_list.append(profit_on_share)
            self.BV_list.append(BV)
            self.cash_on_share_list.append(cash_on_share)
            self.gross_margin_list.append(gross_margin)
            self.ROE_list.append(ROE)
            self.price_sales_on_share_list.append(price_sales_on_share)
            self.P_E_list.append(P_E)
            self.P_BV_list.append(P_BV)
            self.enterprise_value_list.append(enterprise_value)
            self.cash_flow_on_share_list.append(cash_flow_on_share)
            self.dividends_on_share_list.append(dividends_on_share)

        self.overview_of_companies['Jméno'] = self.list_of_companies
        self.overview_of_companies['Poslední obchod'] = self.last_trade_list
        self.overview_of_companies['Cena akcie'] = self.price_list
        self.overview_of_companies['Růst/Pokles'] = self.change_list
        self.overview_of_companies['Objem obchodů'] = self.equity_list
        self.overview_of_companies['Tržní kapitalizace'] = self.market_capitalization_list
        self.overview_of_companies['Celkové tržby (poslední rok)'] = self.total_turnover_list
        self.overview_of_companies['EBITDA (poslední rok)'] = self.EBITDA_list
        self.overview_of_companies['Čistý zisk pro akcionáře (poslední rok)'] = self.pure_gain_for_shareholders_list
        self.overview_of_companies['Zisk na akcii (EPS, poslední rok)'] = self.gain_on_share_list
        self.overview_of_companies['Tržby na akcii (poslední rok)'] = self.profit_on_share_list
        self.overview_of_companies['Účetní hodnota na akcii (BV, poslední rok)'] = self.BV_list
        self.overview_of_companies['Hotovost na akcii (poslední rok)	'] = self.cash_flow_on_share_list
        self.overview_of_companies['Hrubá marže (poslední rok)'] = self.gross_margin_list
        self.overview_of_companies['Návratnost vlastního jmění (ROE, poslední rok)'] = self.ROE_list
        self.overview_of_companies['Cena/tržby na akcii (poslední rok)'] = self.price_sales_on_share_list
        self.overview_of_companies['P/E bez mimořádných položek (poslední rok)'] = self.P_E_list
        self.overview_of_companies['P/BV (poslední rok)'] = self.P_BV_list
        self.overview_of_companies['Cena Enterprise'] = self.enterprise_value_list
        self.overview_of_companies['Cashflow na akcii'] = self.profit_on_share_list
        self.overview_of_companies['Dividenda na akcii (poslední rok)'] = self.dividends_on_share_list

    def pandas_to_excel(self, path):
        # loads scraped data into excel files for later use
        self.overview_of_companies.to_excel(path,
                                            sheet_name='Prehled firem')

    def driver_shutdown(self):
        # shuts down the driver
        self.driver.close()
        print('Your scraping is complete!')


RoboticProcessAutomation(5, 'C:\\Users\\jonas\\PycharmProjects\\RPA_financni_informace\\Informace_o_firmach.xls')

# first argument is the number of companies whose information will be scraped into the output table, second argument is path to output xls file with the table
