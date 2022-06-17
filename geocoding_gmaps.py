'''
Created on Mar 9, 2022

@author: Private
'''
'''
Created on Mar 3, 2022

@author: Private
'''
# to be run after geocoding_nominatim

import pandas as pd
import numpy as np
import time
import requests
from pprint import pprint
from geopy.geocoders import GoogleV3

print('Please ensure a full address is present under a column titled "Full Address", and lat/long from Nominatim in columns titled "Latitude" and "Longitude".')
input_workbook = input('What is the name of the workbook you want to process (with file extension)? ')
state = input('Enter Oregon/Washington: ')

if state == 'Oregon':
    state_alt = 'OR'
else:
    state_alt = 'WA'

input_sheet = input('What is the name of the sheet your data resides on? ')
api_key = input('Enter your API key: ')

df = pd.read_excel(input_workbook, sheet_name=input_sheet)
df1 = df.where(pd.notnull(df), None) #replaces all NaN with None
outputarr = [[]]

geolocator = GoogleV3(api_key=api_key)

row = 0

#iterate over addresses in rows
for ind in range(0, max(df1.index) + 1):
    
    address = df1['Full Address'][ind]
    
    print(address)
    
    if address != None and df1['Latitude'][ind] == None: # we only want to process lines that have an address and that do not yet have a lat/long
        try:
            location = geolocator.geocode(query = address, sensor = False)
            print(location)
        except Exception as e:
            print(e)
        
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
outputdf.to_excel("googlev3.xlsx")
