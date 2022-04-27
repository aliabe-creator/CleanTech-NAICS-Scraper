'''
Created on Sep 16, 2021

@author: Private
'''

import pandas as pd
from xlwt import Workbook
import requests

df = pd.read_excel('C:/Users/abe98/Downloads/test.xlsx', sheet_name='Sheet1') #read the excel file with company urls
df1 = df.where(pd.notnull(df), None)
wb = Workbook() #output workbook
sheet1 = wb.add_sheet("output") #add sheet to output workbook
row = 0

#textrazor api key
headers = {
    'x-textrazor-key': 'afb994c2c671f92c276ce4609be22eb9e54cd5211386696a97a8368d',
}

for ind in df1.index: #iterate over addresses in rows
    if df1['Website address'][ind] != None: #makes sure that a url is actually in the cell
        url = "https://" + df1['Website address'][ind] #formats url
        
        data = {
            'extractors': 'entities',
            'url': url,
            'cleanup.mode': 'cleanHTML'
            }
        response = requests.post('https://api.textrazor.com/', headers=headers, data=data).json() #sending to textrazor
        entities = response['response']['entities'] #using entities because more concise. Topics does provide more, well, topics, but too many.
        
        result = {}
        
        for entity in entities:
            try:
                type = entity['type']
            except:
                type = []
            if entity['confidenceScore'] > 7 and 'Place' not in type and 'Company' not in type: #0.5 - 10, but can go above 10. Filter out places and company names.
                result[str(entity['entityId'])] = entity['confidenceScore']
        
        sheet1.write(row, 0, str(result)) #write to excel file
        print(result)
        
    row = row + 1 #put outside so even if blank, then will properly list

wb.save('xlwt.xls')