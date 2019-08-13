import pandas as pd
from selenium import webdriver
import shutil

import time

import os


class Downloads:
    """Class responsible for downloading report files from web sources"""

    def set_Chrome_driver(self,path):
        """Set driver and driver's options"""

        options = webdriver.ChromeOptions()

        prefs = {
            "download.default_directory": r"C:\Users\reports\PycharmProjects\LOANDO reporting\DOWNLOADS" + path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }

        options.add_experimental_option('prefs', prefs)

        chromedriver = "C:\Selenium Chromedriver\chromedriver.exe"

        self.driver = webdriver.Chrome(executable_path=chromedriver, options=options)


    def clear_downloads(self,path):
        """Clear downloaded file according to provided path"""
        try:
            shutil.rmtree(path)
        except:
            pass

    def login_CRM(self):
        """Logging into CRM"""

        self.driver.get("https://epspl.wcard.int/MemberShipLogin.aspx")

        self.driver.find_element_by_id("Login1_UserName").send_keys(os.environ.get('ERP_USER'))

        self.driver.find_element_by_id("Login1_Password").send_keys(os.environ.get('ERP_PASSWORD'))

        self.driver.find_element_by_id("Login1_LoginButton").click()



    def download_Proposals_DC1(self,code):
        """Download Card Proposals"""

        self.set_Chrome_driver("\Proposals_DC1")

        self.clear_downloads("DOWNLOADS\Proposals_DC1")

        self.login_CRM()

        self.driver.execute_script("aspxGVScheduleCommand('ctl00_ContentPlaceHolder1_gvProposals',['ClearFilter'],0)")

        time.sleep(2)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_gvProposals_DXFREditorcol12_I").send_keys(code)

        time.sleep(5)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_bExport").click()

        time.sleep(10)

        self.driver.close()


    def download_Proposals_DC2(self,code):
        """Download Card Proposals"""

        self.set_Chrome_driver("\Proposals_DC2")

        self.clear_downloads("DOWNLOADS\Proposals_DC2")

        self.login_CRM()

        self.driver.execute_script("aspxGVScheduleCommand('ctl00_ContentPlaceHolder1_gvProposals',['ClearFilter'],0)")

        time.sleep(2)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_gvProposals_DXFREditorcol13_I").send_keys(code)

        time.sleep(5)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_bExport").click()

        time.sleep(10)

        self.driver.close()






    def download_Credit_Cards_DC1(self,code):
        """Download Credit Cards"""

        self.set_Chrome_driver("\CreditCards_DC1")

        self.clear_downloads("DOWNLOADS\CreditCards_DC1")

        self.login_CRM()

        self.driver.get("https://epspl.wcard.int/CreditCards/CreditCards.aspx")

        time.sleep(1)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_ASPxGridViewCreditCards_DXFREditorcol12_I").send_keys(code)

        time.sleep(5)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_bExport").click()

        time.sleep(10)

        self.driver.close()


    def download_Credit_Cards_DC2(self,code):
        """Download Credit Cards"""

        self.set_Chrome_driver("\CreditCards_DC2")

        self.clear_downloads("DOWNLOADS\CreditCards_DC2")

        self.login_CRM()

        self.driver.get("https://epspl.wcard.int/CreditCards/CreditCards.aspx")

        time.sleep(1)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_ASPxGridViewCreditCards_DXFREditorcol13_I").send_keys(code)

        time.sleep(5)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_bExport").click()

        time.sleep(10)

        self.driver.close()


    def concat_Credit_Cards(self):

        df_1 = pd.read_excel("DOWNLOADS\CreditCards_DC1\ASPxGridViewCreditCards.xls")
        df_2 = pd.read_excel("DOWNLOADS\CreditCards_DC2\ASPxGridViewCreditCards.xls")

        df_1 = pd.concat([df_1, df_2], ignore_index=True, sort=False)

        df_1.to_excel("DOWNLOADS\CreditCards\ASPxGridViewCreditCards.xls", index=False)


    def concat_Proposals(self):

        df_1 = pd.read_excel("DOWNLOADS\Proposals_DC1\gvProposals.xls")
        df_2 = pd.read_excel("DOWNLOADS\Proposals_DC2\gvProposals.xls")

        df_1 = pd.concat([df_1, df_2], ignore_index=True, sort=False)

        df_1.to_excel("DOWNLOADS\Proposals\gvProposals.xls", index=False)



    def download_Reports_Cards(self):
        """Download Reports Cards"""

        self.set_Chrome_driver("\ReportsCards")

        self.clear_downloads("DOWNLOADS\ReportsCards")

        self.login_CRM()

        self.driver.get("https://epspl.wcard.int/Reports/AllCards.aspx")

        time.sleep(1)

        self.driver.find_element_by_id("ctl00_ContentPlaceHolder1_bExport").click()

        time.sleep(10)

        self.driver.close()

    def download_Raport_do_CC(self):
        """Copy Raport do CC file"""
        try:
            shutil.copyfile("J:/Public/Karty/Raport do CC NEW.ods",
                        "C:/Users/reports/PycharmProjects/LOANDO reporting/DOWNLOADS/Raport do CC/Raport do CC NEW.ods")
        except:
            print("Raport do CC NEW.ODS WAS NOT FOUND IN THE LOCATION")

    def download_Processing(self):
        """Copy Processing file"""
        try:
            shutil.copyfile("J:/Public/OFFLINE/processing LOANDO.xlsx",
                        "C:/Users/reports/PycharmProjects/LOANDO reporting/DOWNLOADS/Processing/processing LOANDO.xlsx")
        except:
            print("processing LOANDO.XLSX WAS NOT FOUND IN THE LOCATION")




