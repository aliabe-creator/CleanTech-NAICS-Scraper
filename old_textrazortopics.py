'''
Created on Sep 16, 2021

@author: Private
'''

import pandas as pd
from xlwt import Workbook
import requests
from pprint import pprint

df = pd.read_excel('C:/Users/abe98/Downloads/test.xlsx', sheet_name='Sheet1') #read the excel file with company urls
df1 = df.where(pd.notnull(df), None) #replaces all NaN with None
wb = Workbook() #output workbook
sheet1 = wb.add_sheet("output") #add sheet to output workbook
row = 0

#textrazor api key
headers = {
    'x-textrazor-key': 'afb994c2c671f92c276ce4609be22eb9e54cd5211386696a97a8368d',
}

desiredWords = ['clean', 'sustainab', 'renewab', 'environment', 'climate', 'atmosph']

for ind in df1.index: #iterate over addresses in rows
    desScore = 0
    
    if df1['Website address'][ind] != None: #makes sure that a url is actually in the cell
        url = "https://" + df1['Website address'][ind] #formats url
        
        data = {
            'extractors': 'topics',
            'url': url,
            'cleanup.mode': 'cleanHTML'
            }
        response = requests.post('https://api.textrazor.com/', headers=headers, data=data).json() #sending to textrazor, format response as json
        topics = response['response']['topics'] #using entities because more concise. Topics does provide more, well, topics, but too many.
        
        result = {} #create blank dict
        
        for topic in topics:
            if topic['score'] > 0.5: #if the topic is reasonably confident
                result[str(topic['label'])] = topic['score'] #saves in the format 'Energy': 1
                
                splitTopic = topic['label'].split() #split the topic, if multiple words, into separate words in an array
                
                for word in splitTopic: #iterate over each word in split array
                    for desiredWord in desiredWords:
                        if desiredWord in word.lower(): #check to see if desired word is in the lowercase topic word
                            desScore = desScore + 1
        
        sheet1.write(row, 0, str(result)) #write to excel file
        sheet1.write(row, 1, desScore)
        print(result)
        
    row = row + 1 #put outside so even if blank, then will properly list

wb.save('xlwt.xls')