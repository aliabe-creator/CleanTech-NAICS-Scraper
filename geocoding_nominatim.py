'''
Created on Mar 3, 2022

@author: Private
'''

import pandas as pd
import numpy as np
import time
import requests
from pprint import pprint
from geopy.geocoders import Nominatim

print('Please ensure a full address is present under a column titled "Full Address"')
input_workbook = input('What is the name of the workbook you want to process (with file extension)? ')
state = input('Enter Oregon/Washington: ')

if state == 'Oregon':
    state_alt = 'OR'
else:
    state_alt = 'WA'

input_sheet = input('What is the name of the sheet your data resides on? ')

df = pd.read_excel(input_workbook, sheet_name=input_sheet)
df1 = df.where(pd.notnull(df), None) #replaces all NaN with None
outputarr = [[]]

geolocator = Nominatim(user_agent='cleantech')

row = 0

#iterate over addresses in rows
for ind in range(0, max(df1.index) + 1):
    
    address = df1['Full Address'][ind]
    
    print(address)
    
    if address != None:
        location = geolocator.geocode(query = address, country_codes = 'us')
        print(location)
        
        if location != None and (state in str(location) or state_alt in str(location)):
            intermed_array = [location.latitude, location.longitude]
            outputarr.append(intermed_array)
        else:
            outputarr.append('')
    else:
        outputarr.append('')
        
    time.sleep(1)
    row = row + 1 #put outside so even if blank, then will properly list

print(outputarr)
outputdf = pd.DataFrame(outputarr)
outputdf.to_excel("nominatim.xlsx")
