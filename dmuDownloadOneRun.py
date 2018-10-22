
# coding: utf-8

# In[176]:

from selenium import webdriver 
from selenium.common.exceptions import NoSuchElementException 
#from selenium.webdriver.common.by import By 
#from selenium.webdriver.support.ui import WebDriverWait 
#from selenium.webdriver.support import expected_conditions as EC 
#from selenium.common.exceptions import TimeoutException
import time
#import numpy as np
import pandas as pd
import os
from bs4 import BeautifulSoup

os.chdir('\\\\ad.ing.net\\WPS\\NL\\P\\GD\\012223\\01 MARKETING INTELLIGENCE\\Data\\WalletSizeDatabase')

#%%
def check_exists_by_xpath(xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

# In[234]:

#set up chrome environment and driver
option = webdriver.ChromeOptions()
option.add_experimental_option("prefs", {
  "download.default_directory": r"\\ad.ing.net\WPS\NL\P\GD\012223\01 MARKETING INTELLIGENCE\Data\WalletSizeDatabase",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})
#option.add_argument("-incognito")

browser = webdriver.Chrome(executable_path="C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe", chrome_options=option)


# In[235]:

#define urls, usernames and passwords
bvd_home = "https://walletsizingcatalyst.bvdinfo.com/version-2018614/Home.serv?product=WalletSizingcatalyst"
username = "username"
password = "password"


# In[236]:
#open homepage
browser.get(bvd_home)
    
    
# In[237]:
    
    
#login to website
browser.find_element_by_id('inputUser').send_keys(username)
browser.find_element_by_id('inputPassword').send_keys(password)
browser.find_element_by_id('bnLoginNeo').click()
    
  # In[238]:

#Create search list
walletsING = pd.read_excel('\\\\ad.ing.net\\WPS\\NL\\P\\GD\\012223\\01 MARKETING INTELLIGENCE\\Data\\WalletSizeDatabase\\dmuSearchList.xlsx')
walletsING.columns = ['BvDID']
searchBvD = walletsING['BvDID'].to_frame()

#%%
insideDMUDataframeALL = pd.DataFrame(columns=['Company name', 'BvD ID', 'Ctry code', 'Identified DMU name', 'BvD ID', 'Ctry code'])
insideDMUDataframeALL.empty
outsideDMUDataframeALL = pd.DataFrame(columns=['Company name', 'BvD ID', 'Ctry code', 'Identified DMU name', 'BvD ID', 'Ctry code'])
outsideDMUDataframeALL.empty

#%%

#open advance search options
time.sleep(1)
browser.find_element_by_id('ContentContainer_ctl00_Header_ctl00_OpenSearch').click()
window_before = browser.window_handles[0]


#select 5k at once
thisSearchBVD = searchBvD.iloc[5000:,:]
thisSearchBVD['BvDID'] = thisSearchBVD['BvDID'].astype(str) + ','
thisSearchBVD = thisSearchBVD.T
thisSearchBVD = thisSearchBVD.to_string(index=False, header =False)


#Open search by BvD ID
time.sleep(1)
browser.find_element_by_xpath('//*[@id="ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_Menu1"]/li[2]/span').click()
time.sleep(2)
browser.find_element_by_xpath('//*[@id="ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_Menu1"]/li[2]/ul/li[1]/span').click()

#fill the search criteria
time.sleep(3)
str1 = "document.getElementById('MasterContent_ctl00_Content_CodeList').value='"
str2 = thisSearchBVD
str3 = "'"
browser.execute_script(str1+str2+str3)
browser.find_element_by_xpath('//*[@id="MasterContent_ctl00_Content_CodeList"]').send_keys(" ")
time.sleep(3)
browser.find_element_by_id('MasterContent_ctl00_Content_SelectCodeList').click()

#handle outdated ID's
time.sleep(3)
if check_exists_by_xpath('//*[@id="MasterContent_ctl00_Content_IdConversionPopupContent"]') :
    print('outdated')
    browser.find_element_by_id('MasterContent_ctl00_Content_bnOk').click()

#Make search complete and progress
time.sleep(2)
browser.find_element_by_xpath('//*[@id="MasterContent_ctl00_Content_Ok"]/img').click()
time.sleep(2)
browser.find_element_by_id('ContentContainer1_ctl00_Content_QuickSearch1_ctl05_GoToList').click()           

#Start wallet aggregation
time.sleep(2)
browser.find_element_by_id('ContentContainer1_ctl00_Content_ExportButtons_AggregationWizard').click()
time.sleep(4)
browser.find_element_by_id('ContentContainer_ctl00_Content_ctl00_ManuelDMUReplace').click()
time.sleep(2)
browser.find_element_by_css_selector('#ContentContainer_ctl00_Content_NextAjaxPanel > span > img').click()

#Extract table
time.sleep(40) 
datalist = []

#%% Extract table


if check_exists_by_xpath('//*[@id="InsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]') :
    insidePages = browser.find_element_by_xpath('//*[@id="InsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]')
    insidePages = int(insidePages.text.split("Page 1 of ",1)[1])
    
    for i in range(0,insidePages) :
        #Get insideDMU
        #Handover page from selenium to beautifulsoup
        soup_level2=BeautifulSoup(browser.page_source, 'lxml')
        #Beautiful Soup grabs the HTML table on the page
        table = soup_level2.find_all('table')
        #Giving the HTML table to pandas to put in a dataframe object
        df = pd.read_html(str(table),header=0)
        #Store the dataframe in a list
        datalist = []
        datalist.append(df[0])
        
        time.sleep(2)
        
        insideDMUDataframe = df[9]
        insideDMUDataframe.columns = insideDMUDataframe.iloc[1,:]
        insideDMUDataframe = insideDMUDataframe.iloc[2:insideDMUDataframe.shape[0]-1,0:6]
        notEmptyInside = insideDMUDataframe.iloc[:,0].notnull()
        insideDMUDataframe = insideDMUDataframe[notEmptyInside]
        insideDMUDataframeALL = insideDMUDataframeALL.append(insideDMUDataframe)
        browser.find_element_by_xpath('//*[@id="InsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[4]/a/img').click()

#if check_exists_by_xpath('//*[@id="OutsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]') :
#    outsidePages = browser.find_element_by_xpath('//*[@id="OutsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]')
#    outsidePages = int(outsidePages.text.split("Page 1 of ",1)[1])
#    
#    #Get outsideDMU
#    outsideDMUDataframe = df[11]
#    outsideDMUDataframe.columns = outsideDMUDataframe.iloc[1,:]
#    outsideDMUDataframe = outsideDMUDataframe.iloc[2:outsideDMUDataframe.shape[0]-1,0:6]
#    notEmptyOutside = outsideDMUDataframe.iloc[:,0].notnull()
#    outsideDMUDataframe = notEmptyOutside[notEmptyInside]

if check_exists_by_xpath('//*[@id="OutsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]') :
    outsidePages = browser.find_element_by_xpath('//*[@id="OutsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]')
    outsidePages = int(outsidePages.text.split("Page 1 of ",1)[1])
    
    for i in range(0,outsidePages) :
        #Get insideDMU
        #Handover page from selenium to beautifulsoup
        soup_level2=BeautifulSoup(browser.page_source, 'lxml')
        #Beautiful Soup grabs the HTML table on the page
        table = soup_level2.find_all('table')
        #Giving the HTML table to pandas to put in a dataframe object
        df = pd.read_html(str(table),header=0)
        #Store the dataframe in a list
        datalist = []
        datalist.append(df[0])
        
        time.sleep(2)
        
        outsideDMUDataframe = df[11]
        outsideDMUDataframe.columns = outsideDMUDataframe.iloc[1,:]
        outsideDMUDataframe = outsideDMUDataframe.iloc[2:outsideDMUDataframe.shape[0]-1,0:6]
        notEmptyInside = outsideDMUDataframe.iloc[:,0].notnull()
        outsideDMUDataframe = outsideDMUDataframe[notEmptyInside]
        outsideDMUDataframeALL = outsideDMUDataframeALL.append(outsideDMUDataframe)
        browser.find_element_by_xpath('//*[@id="OutsideSelectionDiv"]/table/tbody/tr[1]/td/table/tbody/tr/td[4]/a/img').click()

#%%
#insideDMUDataframeALL_1 = insideDMUDataframeALL
#outsideDMUDataframeALL_1 = outsideDMUDataframeALL
insideDMUDataframeALL.to_excel('\\\\ad.ing.net\\WPS\\NL\\P\\GD\\012223\\01 MARKETING INTELLIGENCE\\Data\\WalletSizeDatabase\\insideDMUDataframeALL_2.xlsx')
outsideDMUDataframeALL.to_excel('\\\\ad.ing.net\\WPS\\NL\\P\\GD\\012223\\01 MARKETING INTELLIGENCE\\Data\\WalletSizeDatabase\\outsideDMUDataframeALL_2.xlsx')
