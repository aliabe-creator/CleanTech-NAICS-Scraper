'''
Created on Feb 25, 2022

@author: Private
'''

import pandas as pd
import numpy as np
import re
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys 
import time
import pyautogui
import pyperclip
import random

input_workbook = input('What is the name of the workbook you want to process (with file extension)? ')
input_sheet = input('What is the name of the sheet your data resides on? ')

df = pd.read_excel(input_workbook, sheet_name=input_sheet)
df1 = df.where(pd.notnull(df), None) #replaces all NaN with None
outputarr = []

state = input('What state is this data from: ')

#start selenium
options = Options();
options.preferences.update({"javascript.enabled": False, "extensions.pocket.enabled": False, "browser.display.show_image_placeholders": False, "browser.display.use_document_fonts": 0, "media.volume_scale": 0}) #disabling JS bypasses captcha
driver = webdriver.Firefox(options=options)
driver.set_page_load_timeout(10)
driver.maximize_window()

row = 0

#iterate over addresses in rows
for ind in range(0, max(df1.index)):
    
    driver.get('https://mailinglists.com/mailinglistsxpress/duns-number-sic-and-naics-code-lookup/') #visit website
    
    #extract relevant info from df from Excel
    company_name = df1['Company Name'][ind]
    city = df1['City'][ind]
    
    #fill out form
    driver.find_element_by_name('naics_lookup[companyName]').send_keys(company_name)
    driver.find_element_by_name('naics_lookup[city]').send_keys(city)
    driver.find_element_by_name('naics_lookup[state]').send_keys(state)
    
    time.sleep(random.uniform(0, 0.5))
    
    #submit form
    driver.find_element_by_name('submit_naics_search').click()
    
    pyautogui.click(x=1, y=200, button = 'right')
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c') #copy text
    
    pagetext=pyperclip.paste() #paste text into variable
    
    #check and make sure NAICS present
    if 'No results found. Please try again.' not in pagetext:
        partitioned = pagetext.partition('NAICS 1 Code')
        following_no_leading = re.sub(r"^\s+" , "" , partitioned[2])
        NAICS = following_no_leading.split()[0]
        print(f'NAICS code found: {NAICS}')
        outputarr.append(NAICS)
    else: #meaning no matching company found
        outputarr.append('')
    
    print(f'{round(row / max(df1.index) * 100, 2)} % done.')
    
    row = row + 1 #put outside so even if blank, then will properly list

print(outputarr)
outputdf = pd.DataFrame(outputarr).T
outputdf.to_excel("naics.xlsx")