
# coding: utf-8

# In[176]:


from selenium import webdriver 
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import TimeoutException
import time
import numpy as np
import pandas as pd


# In[233]:

wallets = 829398
steps = wallets / 10000
steps = steps+1

#create panda
z = (steps,2)
boundriesDF = pd.DataFrame(np.zeros(z))


for i in range(0,steps) :
    boundriesDF.iloc[i,0] = i*10000+1
    boundriesDF.iloc[i,1] =(i+1)*10000
        
y = (steps,3)
hitsDF = pd.DataFrame(np.zeros(y))
hitsDF.iloc[:,0] = boundriesDF.iloc[:,0]
hitsDF.iloc[:,1] = boundriesDF.iloc[:,1]

# In[234]:


#set up chrome environment and driver
option = webdriver.ChromeOptions()
option.add_experimental_option("prefs", {
  "download.default_directory": r"\\ad.ing.net\WPS\NL\P\GD\012223\01 MARKETING INTELLIGENCE\Data\scrapeBvD",
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

#open advance search options
time.sleep(1)
browser.find_element_by_id('ContentContainer_ctl00_Header_ctl00_OpenSearch').click()
window_before = browser.window_handles[0]
  
# In[239]:


#Go to the opperating revenue search
browser.find_element_by_xpath("//ul[@id='ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_Menu2']/li/span").click()
time.sleep(1)
browser.find_element_by_xpath("//ul[@id='ContentContainer1_ctl00_Content_QuickSearch1_ctl02_SearchSearchMenu_Menu2']/li/ul/li[3]/span").click()

# In[240]:


#Enter search criteria
minSearch = 10000

browser.find_element_by_id('MasterContent_ctl00_Content_MatrixControl_MinFreeText').send_keys(str(minSearch))
#browser.find_element_by_id('MasterContent_ctl00_Content_MatrixControl_MaxFreeText').send_keys(str(upperBound))
browser.find_element_by_css_selector('#MasterContent_ctl00_Content_MatrixControl_Ok > img').click()


# In[243]:

#go to the list of results
browser.find_element_by_id('ContentContainer1_ctl00_Content_QuickSearch1_ctl05_GoToList').click()

#Get CMI wallet
browser.find_element_by_css_selector("#ContentContainer1_ctl00_Content_tabBarControl_tabBarControl_listformatcompanies0001 > div.TabBarBaseItem.TabBarInactiveItem > span.TabBarItemLabel").click()
window_before = browser.window_handles[0]

for scrapeRUN in range(steps) :    
    

        
    #open export CMI wallet
    browser.find_element_by_id("ContentContainer1_ctl00_Content_ExportButtons_ExportButton").click()
    
    
    #Switch to other window
    window_after = browser.window_handles[1]
    browser.switch_to_window(window_after)
    
    lowerBound = int(boundriesDF.iloc[scrapeRUN,0])
    upperBound = int(boundriesDF.iloc[scrapeRUN,1])
    
    export_name = '_'+'Run'+'_'+str(lowerBound)+'_'+str(upperBound)+'_'
    browser.find_element_by_name('RANGEFROM').click()
    browser.find_element_by_name('RANGEFROM').send_keys(str(lowerBound))
    browser.find_element_by_name('RANGETO').click()
    browser.find_element_by_name('RANGETO').send_keys(str(upperBound))
    
    browser.find_element_by_id("ctl00_ContentContainer1_ctl00_LowerContent_Formatexportoptions1_ExportDisplayName").click()
    browser.find_element_by_id("ctl00_ContentContainer1_ctl00_LowerContent_Formatexportoptions1_ExportDisplayName").send_keys(export_name)
    browser.find_element_by_id("imgBnOk").click()
    time.sleep(180)
    browser.close()
       
    # reselect old page page
    browser.switch_to_window(window_before)
