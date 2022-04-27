'''
Created on Dec 9, 2021

Python program to search for email addresses within an Excel sheet, extract their domains, and check said domain names for existing MX records.

Compatibility: Windows, minor edits required for Linux/MacOS.
'''

#script to go through Excel file and validate domains/email addresses

import pandas as pd
import numpy as np
import re
import subprocess

df = pd.read_excel("wash_large.xlsx", sheet_name='Sheet1')
df1 = df.where(pd.notnull(df), None) #replaces all NaN with None
outputarr = []

print(df1)

#counts what row we are on
row = 0

#iterate over addresses in rows
for ind in range(65536, max(df1.index)):
    print(row)
    if df1['Website address'][ind] != None: #makes sure that a url is actually in the cell
        cmdoutput = subprocess.run(['nslookup', str(df1['Website address'][ind])], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        if 'Non-existent domain' in cmdoutput.stderr.decode('utf-8') or 'timed out' in cmdoutput.stderr.decode('utf-8'):
            outputarr.append('Domain invalid.')
        else:
            outputarr.append('Domain valid.')
    if df1['Website address'][ind] == None: #meaning no url
        outputarr.append('No domain.')
    
    row = row + 1 #put outside so even if blank, then will properly list

print(outputarr)
outputdf = pd.DataFrame(outputarr).T
outputdf.to_excel("C:/Users/abe98/Desktop/sus.xlsx")

#~~Below code is specifically for email addresses~~
'''
oneemailcells = np.array(df['Contact'][df['Contact'].str.contains('@', regex=False)].values.tolist())

def returninvaliddomain(emailcells):
    mxnotpresent = []
    
    stremailcells = str(emailcells)
    
    domains = re.findall("@([a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", stremailcells)
    
    for domain in domains:
        cmdoutput = subprocess.run(['nslookup', '-type=mx', domain], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        if 'MX preference = 0, mail exchanger = (root)' in cmdoutput.stdout.decode('utf-8') or 'Non-existent domain' in cmdoutput.stderr.decode('utf-8'):
            mxnotpresent.append(domain)
            
    print(mxnotpresent)
            
returninvaliddomain(oneemailcells)
'''