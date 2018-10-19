# -*- coding: utf-8 -*-

#Created on Thu Jul  5 14:45:01 2018
#@author: EY36VI

#%%
import pandas as pd
#import numpy as np
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException  
import time
import os

##change working directory
os.chdir('\\\\ad.ing.net\\WPS\\NL\\P\\GD\\012223\\01 MARKETING INTELLIGENCE\\Data\\scrapeOrbis')

#%% set up chrome environment and driver

option = webdriver.ChromeOptions()
option.add_experimental_option("prefs", {
  "download.default_directory": r"\\ad.ing.net\WPS\NL\P\GD\012223\01 MARKETING INTELLIGENCE\Data\scrapeOrbis",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})
option.add_argument("-incognito")

#setting of device important for driver location (option:'own' | 'remote')
comp_setting = 'own' 

if comp_setting == 'own' :
    browser = webdriver.Chrome(executable_path="C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe", chrome_options=option)
elif comp_setting == 'remote' :
    browser = webdriver.Chrome(executable_path="C:\Anaconda\chromedriver.exe", chrome_options=option)

#%% Functions
def check_exists_by_css(css):
    try:
        browser.find_element_by_css_selector(css)
    except NoSuchElementException:
        return False
    return True

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

#%% Login to Orbis database

#define urls, usernames and passwords
bvd_home = "https://orbis4.bvdinfo.com/version-201874/orbis/1/Companies/Login?returnUrl=%2Fversion-201874%2Forbis%2F1%2FCompanies"
username = "username"
password = "password"

#open homepage
browser.get(bvd_home)

##login to website
browser.find_element_by_id('user').send_keys(username)
browser.find_element_by_id('pw').send_keys(password)
browser.find_element_by_css_selector('button.ok').click()

time.sleep(1)
window_start = browser.window_handles[0]

#%% loop through companies to find corperate structure
#%%
#import eups to bvd we will have to search
targetEUPS = pd.read_excel('SegClients.xlsx', sheet_name  = 'MappedClients')

#Boundries
beginEUP = 191
endEUP = 4000
#endEUP = targetEUPS.shape[0]

for downloadTick in range(beginEUP,endEUP) :
    print('Current company --- ' + str(downloadTick))
    print(targetEUPS.loc[downloadTick])
    companyId = targetEUPS.iloc[downloadTick,6]
    companyEUP = targetEUPS.iloc[downloadTick,2]
    companyName = targetEUPS.iloc[downloadTick,1]
    
    #Open company book
    if check_exists_by_id('search') :
        browser.find_element_by_id('search').send_keys(companyId)
    else :
        time.sleep(10)
        browser.find_element_by_id('search').send_keys(companyId)
    
    time.sleep(4)
    if check_exists_by_css('p.name') :
        browser.find_element_by_css_selector('p.name').click()
                                           
        #Go to corporate structure page in book
        time.sleep(5)
        if check_exists_by_css("li.midgetHolder.Ownership > a.showMore") :
            browser.find_element_by_css_selector("li.midgetHolder.Ownership > a.showMore").click()
        else :
            time.sleep(10)
            browser.find_element_by_css_selector("li.midgetHolder.Ownership > a.showMore").click()
        
        #Open the full company structure
        time.sleep(15)
        if check_exists_by_xpath("(//a[contains(text(),'Add/remove columns')])[2]") :
            #retrive all relevant data from corp structure
            browser.find_element_by_xpath("(//a[contains(text(),'Add/remove columns')])[2]").click()
            time.sleep(3)
            browser.find_element_by_css_selector('div[title="Company info"]').click()
            time.sleep(1)
            browser.find_element_by_id('SUBSIDIARIES*SUBSIDIARIES.-9305:UNIVERSAL').click()
            time.sleep(1)
            browser.find_element_by_css_selector('div[title="Financials"]').click()
            time.sleep(1)
            browser.find_element_by_id('SUBSIDIARIES*SUBSIDIARIES.-9303:UNIVERSAL').click()
            time.sleep(1)
            browser.find_element_by_css_selector('input.button.ok').click()
            time.sleep(8)
            if check_exists_by_css("div.current-unfoldLevel > a > img.px10") :
                browser.find_element_by_css_selector("div.current-unfoldLevel > a > img.px10").click()
                browser.find_element_by_xpath('//li[10]/label').click()
                browser.find_element_by_css_selector('li.close.autoUnfold').click()
            else :
                time.sleep(10)
                browser.find_element_by_css_selector("div.current-unfoldLevel > a > img.px10").click()
                browser.find_element_by_xpath('//li[10]/label').click()
                browser.find_element_by_css_selector('li.close.autoUnfold').click()
                time.sleep(10)
                    
            # Extract data from form
            time.sleep(15)
            soup_level2 = BeautifulSoup(browser.page_source, 'lxml')
            form = soup_level2.find_all('form')[0]
            df = pd.read_html(str(form),header=0)   
            benchMark = 0
            for check in range(len(df)) :
                if df[check].shape[0] > 2 and df[check].shape[1] == 10 :
                    if df[check].iloc[2,0].find('Global Ultimate Owner') != -1 :
                        dfCount = check
            companyDF = pd.DataFrame(df[dfCount])
            names = list(companyDF.columns.values)
            del companyDF[names[0]]
            companyDF = companyDF.drop(companyDF.index[[0,1,2,4,5,6]])
            companyDF.columns = ['Name', 'Country','DirectOwn','TotalOwn','CorpLvl','Source','Timestamp','BVD_ID','OpRevenue']
            companyDF = companyDF[companyDF.Name.str.contains('This company has some subsidiaries but none of them are ultimately owned') == False]
            companyDF['EUP'] = companyEUP
            companyDF['GUO'] = companyDF.iloc[0,7]
            companyDF = companyDF.drop_duplicates()
            
        elif check_exists_by_xpath("//a[contains(text(),'Add/remove columns')]"):
            #If there is no corp structure we take current company as eup/guo etc.
            companyDF = pd.DataFrame(index=range(0,1),columns=['Name', 'Country','DirectOwn','TotalOwn','CorpLvl','Source','Timestamp','BVD_ID','OpRevenue'])
            companyDF['Name'] = companyName
            companyDF['CorpLvl'] = 1
            companyDF['BVD_ID'] = companyId
            companyDF['EUP'] = companyEUP
            companyDF['GUO'] = companyId
            
        elif not check_exists_by_xpath("(//a[contains(text(),'Add/remove columns')])[2]") :
            #Very large corporates have to have extra waiting time
            time.sleep(200)
            if check_exists_by_xpath("(//a[contains(text(),'Add/remove columns')])[2]") :       
                browser.find_element_by_xpath("(//a[contains(text(),'Add/remove columns')])[2]").click()
                time.sleep(3)
                browser.find_element_by_css_selector('div[title="Company info"]').click()
                time.sleep(1)
                browser.find_element_by_id('SUBSIDIARIES*SUBSIDIARIES.-9305:UNIVERSAL').click()
                time.sleep(1)
                browser.find_element_by_css_selector('div[title="Financials"]').click()
                time.sleep(1)
                browser.find_element_by_id('SUBSIDIARIES*SUBSIDIARIES.-9303:UNIVERSAL').click()
                time.sleep(1)
                browser.find_element_by_css_selector('input.button.ok').click()
                time.sleep(150)
                browser.find_element_by_css_selector("div.current-unfoldLevel > a > img.px10").click()
                browser.find_element_by_xpath('//li[10]/label').click()
                browser.find_element_by_css_selector('li.close.autoUnfold').click()
                time.sleep(150)
                
                # Extract data from form
                soup_level2 = BeautifulSoup(browser.page_source, 'lxml')
                form = soup_level2.find_all('form')[0]
                df = pd.read_html(str(form),header=0)   
                benchMark = 0
                for check in range(len(df)) :
                    if df[check].shape[0] > 2 and df[check].shape[1] == 10 :
                        if df[check].iloc[2,0].find('Global Ultimate Owner') != -1 :
                            dfCount = check
                companyDF = pd.DataFrame(df[dfCount])
                names = list(companyDF.columns.values)
                del companyDF[names[0]]
                companyDF = companyDF.drop(companyDF.index[[0,1,2,4,5,6]])
                companyDF.columns = ['Name', 'Country','DirectOwn','TotalOwn','CorpLvl','Source','Timestamp','BVD_ID','OpRevenue']
                companyDF = companyDF[companyDF.Name.str.contains('This company has some subsidiaries but none of them are ultimately owned') == False]
                companyDF['EUP'] = companyEUP
                companyDF['GUO'] = companyDF.iloc[0,7]
                companyDF = companyDF.drop_duplicates()
            
            else :
                #Default end state will get corp lvl of -1
                companyDF = pd.DataFrame(index=range(0,1),columns=['Name', 'Country','DirectOwn','TotalOwn','CorpLvl','Source','Timestamp','BVD_ID','OpRevenue'])
                companyDF['Name'] = companyName
                companyDF['CorpLvl'] = -1
                companyDF['BVD_ID'] = companyId
                companyDF['EUP'] = companyEUP
                companyDF['GUO'] = companyId
    else :
        companyDF = pd.DataFrame(index=range(0,1),columns=['Name', 'Country','DirectOwn','TotalOwn','CorpLvl','Source','Timestamp','BVD_ID','OpRevenue'])
        companyDF['Name'] = companyName
        companyDF['CorpLvl'] = -1
        companyDF['BVD_ID'] = 'NA'
        companyDF['EUP'] = companyEUP
        companyDF['GUO'] = 'NA'
        
    #%% Add to total result
    # initiate dataframe
    if downloadTick == beginEUP :
        BvDCompanyStruc = companyDF
    
    #append new company    
    BvDCompanyStruc = BvDCompanyStruc.append(companyDF, ignore_index=True)
    
    if downloadTick%100 == 0 :
        BvDCompanyStruc.to_pickle('BvDCompanyStruc_intSave')       
        
        
    #%% restart search
    
    #step 1 restart search
    if check_exists_by_css('div.database-selector-main > img.px40') :
        browser.find_element_by_css_selector('div.database-selector-main > img.px40').click()
    else :
        time.sleep(10)
        browser.find_element_by_css_selector('div.database-selector-main > img.px40').click()
    
    #step 2 restart search
    time.sleep(1)
    if check_exists_by_css('li.selected') :
        browser.find_element_by_css_selector('li.selected').click()
    else :
        time.sleep(10)
        browser.find_element_by_css_selector('li.selected').click()
    
    #step 3 restart search
    time.sleep(1)
    if check_exists_by_css('ul.search-buttons.active > li > a') :
        browser.find_element_by_css_selector('ul.search-buttons.active > li > a').click()
    else :
        time.sleep(10)
        browser.find_element_by_css_selector('ul.search-buttons.active > li > a').click()

    #step 4 restart search
    time.sleep(1)
    if check_exists_by_xpath("//a[contains(text(),'Start again')]") :
        browser.find_element_by_xpath("//a[contains(text(),'Start again')]").click()
    else :
        time.sleep(10)
        browser.find_element_by_xpath("//a[contains(text(),'Start again')]").click()

    #sleep before going on
    time.sleep(3)  
    
#%% write pickle and done
stringPickle = 'BvDCompanyStruc_'+ str(beginEUP) + '_' + str(endEUP)
stringExcel = stringPickle + '.xlsx'

BvDCompanyStruc.to_excel(stringExcel) 
BvDCompanyStruc.to_pickle(stringPickle)

print('Done ** last EUP number --- ' + str(downloadTick))
