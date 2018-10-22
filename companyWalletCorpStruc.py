
# coding: utf-8

# In[176]:


from selenium import webdriver 
#from selenium.webdriver.common.by import By 
#from selenium.webdriver.support.ui import WebDriverWait 
#from selenium.webdriver.support import expected_conditions as EC 
#from selenium.common.exceptions import TimeoutException
import time
import os
import numpy as np
import pandas as pd

# In[234]:


#set up chrome environment and driver
option = webdriver.ChromeOptions()
option.add_experimental_option("prefs", {
  "download.default_directory": r"\path\to\folder\",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})
#option.add_argument("-incognito")

#setting of device important for driver location (option:'own' | 'remote')
comp_setting = 'own' 

if comp_setting == 'own' :
    browser = webdriver.Chrome(executable_path="C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe", chrome_options=option)
elif comp_setting == 'remote' :
    browser = webdriver.Chrome(executable_path="C:\Anaconda\chromedriver.exe", chrome_options=option)
    
os.chdir('\path\to\folder\')


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
  
# In[239]:

#Open Saved 
browser.find_element_by_id("ContentContainer1_ctl00_Content_QuickSearch1_ctl02_TabSavedSearches").click()
time.sleep(3)
browser.find_element_by_id("ContentContainer1_ctl00_Content_QuickSearch1_ctl02_MySavedSearches1_DataGridResultViewer_ctl04_Linkbutton1").click()

# In[243]:
window_before = browser.window_handles[0]

numComps = 4279
stepSize = 5
steps = numComps / stepSize
steps = steps+1
#create panda
z = (steps,2)
compBounds = pd.DataFrame(np.zeros(z))
for j in range(0,steps) :
    compBounds.iloc[j,0] = j*stepSize+1
    compBounds.iloc[j,1] =(j+1)*stepSize
compBounds.iloc[-1,1] = numComps   

for scrapeRUN in range(compBounds.size) :           
    #open export functionality
    time.sleep(5)
    browser.find_element_by_id("ContentContainer1_ctl00_Content_ExportButtons_ExportButton").click()
    
    #Switch to other window
    window_after = browser.window_handles[1]
    browser.switch_to_window(window_after)

    browser.find_element_by_xpath('//*[@id="Select1"]/option[6]').click()
    
    lowerBound = compBounds.iloc[scrapeRUN,0]
    upperBound = compBounds.iloc[scrapeRUN,1]
    
    export_name = '_'+'CompSub'+'_'+str(int(lowerBound))+'_'+str(int(upperBound))+'_'
    browser.find_element_by_name('RANGEFROM').click()
    browser.find_element_by_name('RANGEFROM').send_keys(str(lowerBound))
    time.sleep(1)
    browser.find_element_by_name('RANGETO').click()
    browser.find_element_by_name('RANGETO').send_keys(str(upperBound))
    
    time.sleep(1)
    browser.find_element_by_id("ctl00_ContentContainer1_ctl00_LowerContent_Formatexportoptions1_ExportDisplayName").click()
    time.sleep(1)
    browser.find_element_by_id("ctl00_ContentContainer1_ctl00_LowerContent_Formatexportoptions1_ExportDisplayName").send_keys(export_name)
    time.sleep(3)
    browser.find_element_by_id("imgBnOk").click()
    time.sleep(55)
    browser.close()
       
    # reselect old page page
    print(export_name)
    browser.switch_to_window(window_before)
    
