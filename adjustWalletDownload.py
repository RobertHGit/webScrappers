
# coding: utf-8

# In[176]:

from selenium import webdriver 
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from random import randint
import datetime
import time
import numpy as np
import pandas as pd
import os
from bs4 import BeautifulSoup

os.chdir('\path\to\folder\')

#%% partial functions
def check_exists_by_xpath(xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_id(ID):
    try:
        browser.find_element_by_id(ID)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_xpathVisual(xpath) :
    try:
        browser.find_element_by_xpath(xpath)
    except ElementNotVisibleException:
        return False
    return True

def wait_page(name, driver, element_by=By.ID, timeout=15):
    
    ''' 
    This function waits for an element to be loaded - and if that takes too long
    issues a timeout
    '''
    
    try:
        element_present = EC.presence_of_element_located((element_by, name))
        WebDriverWait(driver, timeout).until(element_present)
    except TimeoutException:
        print("Timed out waiting for page to load at " + name)


#%% Search functions

#===================================#
def adjustWalletPath(BvD_ID):
    
    '''
    Path to go to the adjusted wallet
    '''
    
    #enter BvD search id
    wait_page('ContentContainer_ctl00_Content_ctl00_iQuick', driver=browser)
    browser.find_element_by_id('ContentContainer_ctl00_Content_ctl00_iQuick').send_keys(BvD_ID)


    #check to see if BvD ID is valid, and if valid click on company name
    try:
        
        wait_page('a.name', browser, element_by = By.CSS_SELECTOR)
        browser.find_element_by_css_selector('a.name').click()

    #otherwise break this loop
    except NoSuchElementException:
        
        return pd.DataFrame(data = {'bvd_id':[BvD_ID] ,'Wallet':['False']})

    ##Go to adjust wallet
    wait_page('div.divAsTD', driver=browser, element_by = By.CSS_SELECTOR)
    browser.find_element_by_css_selector('div.divAsTD').click()
    
    ## make several steps to get to adjusted wallet
    #Financials
    wait_page('ContentContainer_ctl00_Content_NextAjaxPanel', driver=browser)
    browser.find_element_by_id('ContentContainer_ctl00_Content_NextAjaxPanel').click()
    #Industry
    time.sleep(4)
    wait_page('ContentContainer_ctl00_Content_NextAjaxPanel', driver=browser)
    browser.find_element_by_id('ContentContainer_ctl00_Content_NextAjaxPanel').click()
    #Size
    time.sleep(4)
    wait_page('ContentContainer_ctl00_Content_NextAjaxPanel', driver=browser)
    browser.find_element_by_id('ContentContainer_ctl00_Content_NextAjaxPanel').click()
    #Risk Rating
    time.sleep(4)
    wait_page('ContentContainer_ctl00_Content_NextAjaxPanel', driver=browser)
    browser.find_element_by_id('ContentContainer_ctl00_Content_NextAjaxPanel').click()
    #Pricing
    time.sleep(4)
    wait_page('ContentContainer_ctl00_Content_NextAjaxPanel', driver=browser)
    browser.find_element_by_id('ContentContainer_ctl00_Content_NextAjaxPanel').click()
    
    #Random adjust wallet name (based on time)
    time.sleep(randint(0, 5))
    t = datetime.datetime.now()
    tString = t.strftime('%a, %d %b %Y %H:%M:%S')
    wait_page('WalletName', driver=browser)
    browser.find_element_by_id('WalletName').send_keys(tString)
    browser.find_element_by_id('ContentContainer_ctl00_Content_NextAjaxPanel').click()
    
    #Open whole wallet
    wait_page('a[name="allDetails"]', browser, element_by = By.CSS_SELECTOR)
    browser.find_element_by_css_selector('a[name="allDetails"]').click()
#===================================

def extractTable(browser, TabNum):
    
    '''
    Code will extract all the tables on the page and extract table that has been identified as usefull
    '''
    #Extract html to beautifulsoup
    soup_level2=BeautifulSoup(browser.page_source, 'lxml')
    
    #Beautiful Soup grabs the HTML table on the page
    table = soup_level2.find_all('table')
    
    #Giving the HTML table to pandas to put in a dataframe object
    df = pd.read_html(str(table),header=0)
    df = df[TabNum]
    
    return df
#===================================   

        
# In[234]:

#set up chrome environment and driver
option = webdriver.ChromeOptions()
option.add_experimental_option("prefs", {
  "download.default_directory": r"\path\to\folder\",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})
option.add_argument("-incognito")

browser = webdriver.Chrome(executable_path="C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe", chrome_options=option)


# In[235]:

#define urls, usernames, passwords
bvd_home = "https://walletsizingcatalyst.bvdinfo.com/version-2018614/Home.serv?product=WalletSizingcatalyst"
username = "username"
password = "password"

#start browser
browser.get(bvd_home)

#login to website
browser.find_element_by_id('inputUser').send_keys(username)
browser.find_element_by_id('inputPassword').send_keys(password)
browser.find_element_by_id('bnLoginNeo').click()
   
# In[238]:

#Create search list
walletsING = pd.read_excel('Plat_Ids.xlsx')
adjustedWalletDF = pd.DataFrame(data = {'walletSeg':'empty', 'walletComp':'empty', 'walletType':'empty', 'walletValue':0}, index=[0])
corpDataDF = pd.DataFrame(data = {'inputSeg':'empty', 'inputType':'empty', 'inputValue':0, 'dummy':0}, index=[0])

start = 0
end = 5

for eup in range(start,end) :
    #Random wait
    time.sleep(randint(0, 10))
    
    #get eup bvd id
    eupBvD = walletsING.iloc[eup,0]
    
    #Create adjusted wallet
    adjustWalletPath(eupBvD)
    
    #Extract table [Adjusted wallet]
    TabNumAW = 17 
    adjustedWalletDFRUN = extractTable(browser, TabNumAW)
    NumCut = adjustedWalletDFRUN.index[adjustedWalletDFRUN['Based on the account of'] == 'Total Banking Revenues'].tolist()
    adjustedWalletDFRUN = adjustedWalletDFRUN.iloc[:NumCut[0]+1,:]
    date1 = adjustedWalletDFRUN.columns[2]
    adjustedWalletDFRUN.columns = ['walletSeg', 'walletComp', 'walletType', 'walletValue']
    adjustedWalletDFRUN['eupBvDID'] = eupBvD
    adjustedWalletDFRUN['date'] = date1
    adjustedWalletDF = adjustedWalletDF.append(adjustedWalletDFRUN, ignore_index=True)
    
    #Extract table [input]
    TabNumCD = 18 
    corpDataDFRUN = extractTable(browser, TabNumCD)
    date2 = adjustedWalletDFRUN.columns[3]
    corpDataDFRUN.columns = ['inputSeg', 'inputType', 'inputValue', 'dummy']
    corpDataDFRUN['eupBvDID'] = eupBvD
    corpDataDFRUN['date'] = date2
    corpDataDF = corpDataDF.append(corpDataDFRUN, ignore_index=True)

    #Go back to home page
    wait_page('[title="Return to the home page."]', browser, element_by = By.CSS_SELECTOR)
    browser.find_element_by_css_selector('[title="Return to the home page."]').click()

#%% Save
adjustString = 'adjustedWalletDF_'+str(start)+'_'+str(end-1)
corpDataString = 'corpDataDF_'+str(start)+'_'+str(end-1)

adjustedWalletDF.to_excel(adjustString+'.xlsx')
corpDataDF.to_excel(corpDataString+'.xlsx')