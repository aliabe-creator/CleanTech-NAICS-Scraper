'''
Created on Sep 16, 2021

See attached documentation for detailed description.

Main web crawler program, designed to (1) sequentially visit URLs listed in an Excel column, (2) enumerate sub-links identified on the Excel URLs, (3) compare page text
against a supplied set of regex keys and tally how many hits per site, (4) write number of hits and errors back into Excel file.

Compatibility: Windows/Linux, likely MacOS also supported.
'''

EXCEL_WORKBOOK_NAME = 'baddomains.xlsx'
EXCEL_SHEET_NAME = 'cs_regex_bad_domains'
OUTPUT_WORKBOOK_NAME = 'cs.xls'
OUTPUT_SHEET_NAME = 'output'
STARTING_ROW = 2165 # <-- Do note, pandas starts rows from 0, and the first row is the first non-header row. For example, even if there are 5 header rows, this var should be
                 # 0 to start with the first actual website address.
CONNECTIVITY_CHECK_URL = 'https://brave.com'
WEBSITE_ADDRESS_COLUMN = 'Website address' #string header of the column containing the URLs

#quick config to disable failsafe timeout between pyautogui actions, saves time in the long run.
pyautogui.PAUSE = 0

try:
    import pandas as pd
    from xlwt import Workbook
    from selenium import webdriver
    from selenium.webdriver.firefox.options import Options
    import pyperclip
    import pyautogui
    import re
    import urllib.request
    import time
except ImportError as e:
    print('Please check required Python modules have been installed. Error is included below:\n' + str(e))

#--Scoring helper functions--
def general(s): #returns 1 if any regex in this section matches
    score = 0
    
    regex_terms = ['sustainab', 'green good', 'green technolog', 'green innov', 'eco[^ ]*innov', 'green manufac', 'green prod', 'pollut',
               'ecolabel', 'environ.* product declarat', 'EPD.*environ|environ.*EPD', 'environ.* prefer.* product', 'environ.* label']
    
    for term in regex_terms:
        yesorno = re.search(term, s)
        if yesorno:
            score = 1
    
    if score == 1:
        print('general hit')
            
    return score

def environ(s): #returns +1 if any regex in each section matches
    score = 0
    
    #all-purpose
    if (re.search('natur.* environ', s) or re.search('environ.* friend', s) or re.search('environment.* conserv', s) or re.search('biocompat', s) 
        or re.search('biodivers', s) or re.search('filter', s) or re.search('filtra', s) or re.search('synth.* gas', s) or re.search('regenerat', s)
        or re.search('recircul', s) or re.search('gasification', s) or re.search('gasifier', s) or re.search('fluidized clean gas', s)
        or re.search('gas cleaning', s)):
        
        score += 1
        print('environ - all purpose hit')
    
    #biotreatment
    if ((re.search('biogas.*', s) or re.search('bioreact.*', s) or re.search('polyolef.*', s) or re.search('biopolymer.*', s) 
        or re.search('disinfect.*', s) or re.search('biofilm.*', s) or re.search('biosens.*', s) or re.search('biosolid.*', s) 
        or re.search('caprolact.*', s) or ((re.search('ultraviol.*', s) or re.search('UV', s)) and (re.search('radiat.*', s) 
        or re.search('sol.*', s)))) and (re.search('bioremed.*', s) or re.search('biorecov.*', s) or (re.search('biolog.* treat.*', s)) 
        or re.search('biodegrad.*', s))):
        
        score += 1
        print('environ - biotreatment hit')
    
    #air pollution
    if ((re.search('air.* contr.*', s) or re.search('dust.* contr.*', s) or re.search('particular.* contr.*', s) 
        or re.search('air.* qual.*', s)) and re.search('pollut.*', s)):
        
        score += 1
        print('environ - air pollut hit')
    
    #environmental monitoring
    if  (re.search('environ.* monitor.*', s) and ((re.search('environ.* and instrument.*', s) 
        or re.search('environ.* and analys.*', s)) or re.search('life.*cycle analysis', s) or re.search('life cycle analys.*', s))):
        
        score += 1
        print('environ - monitoring hit')
    
    #marine pollution
    if re.search('marin.* control.*pollut|pollut.*marin.* control', s):
        score += 1
        print('environ - marine pollut hit')
    
    #noise and vib control
    if (re.search('nois.* abat.*', s) or re.search('nois.* reduc.*', s) or re.search('nois.* lessen.*', s)):
        score += 1
        print('environ - noise vib control hit')
    
    #land reclamation
    if (re.search('land', s) and (re.search('reclam.*', s) or re.search('remediat.*', s) or re.search('contamin.*', s))):
        score += 1
        print('environ - land reclamation hit')
        
    #waste management
    if (re.search('wast.*', s) or re.search('sewag.*', s) or re.search('inciner.*', s)):
        score += 1
        print('environ - waste manage hit')
        
    #water supply, treatment
    if ((re.search('slurr.*', s) or re.search('sludg.*', s) or re.search('aque.* solution.*', s) or re.search('wastewat.*', s) 
        or re.search('effluent.*', s) or re.search('sediment.*', s) or re.search('floccul.*', s) or re.search('detergen.*', s) 
        or re.search('coagul.*', s) or re.search('dioxin.*', s) or re.search('flow.* control.* dev.*', s) 
        or re.search('fluid commun.*', s) or re.search('high purit.*', s) or re.search('impur.*', s) or re.search('zeolit.*', s)) 
        and (re.search('water treat.*', s) or re.search('water purif.*', s) or re.search('water pollut.*', s))):
        
        score += 1
        print('environ - water supply hit')
    
    #recovery, recyling
    if (re.search('recycl.*', s) or re.search('compost.*', s) or re.search('stock process.*', s) or re.search('coal combust.*', s) 
        or re.search('remanufactur.*', s) or (re.search('coal', s) and re.search('PCC', s)) 
        or re.search('circulat.* fluid.* bed combust.*', s) or (re.search('combust.*', s) and re.search('CFBC', s))):
        
        score += 1
        print('environ - recycling hit')
        
    return score

def renew(s):
    score = 0
    
    #all-purpose
    if (re.search('renewabl.*', s) and (re.search('energ.*', s) or re.search('electric.*', s))):
        score += 1
        print('renew - all purpose hit')
    
    #wave, tidal
    if ((re.search('two basin schem.*', s) or re.search('wave.* energ.*', s) or re.search('tid.* energ.*', s)) and re.search('electric.*', s)):
        score += 1
        print('renew - wave, tidal hit')
    
    #biomass
    if (re.search('biomass.*', s) or re.search('enzymat.* hydrolys.*', s) or re.search('bio.*bas.* product.*', s)):
        score += 1
        print('renew - biomass hit')
    
    #wind
    if (re.search('wind power.*', s) or re.search('wind energ.*', s) or re.search('wind farm.*', s) or re.search('turbin.* and wind.*', s)):
        score += 1
        print('renew - wind hit')
    
    #geotherm
    if ((re.search('whole system.*', s) and re.search('geotherm.*', s)) or re.search('geotherm.*', s) or re.search('geoexchang.*', s)):
        score += 1
        print('renew - geotherm hit')
    
    #solar
    if (re.search('solar.*', s) and (re.search('ener.*', s) or re.search('linear fresnel sys.*', s) or re.search('electric.*', s) 
        or re.search('cell.*', s) or re.search('heat.*', s) or re.search('cool.*', s) or re.search('photovolt.*', s) 
        or re.search('PV', s) or re.search('cdte', s) or re.search('cadmium tellurid.*', s) or re.search('PVC-U', s) 
        or re.search('photoelectr.*', s) or re.search('photoactiv.*', s) or re.search('sol.*gel.* process.*', s) 
        or re.search('evacuat.* tub.*', s) or re.search('flat plate collect.*', s) or re.search('roof integr.* system.*', s))):
        
        score += 1
        print('renew - solar hit')
    
    return score

def elc(s):
    score = 0
    
    #all purp
    if (re.search('low carbon', s) or re.search('zero carbon', s) or re.search('no carbon', s) or re.search('0 carbon', s) 
        or re.search('low.*carbon', s) or re.search('zero.*carbon', s) or re.search('no.*carbon', s)):
        
        score += 1
        print('elc - all purpose hit')
    
    #alt fuel vehicle
    if (re.search('electric.* vehic.*', s) or re.search('hybrid vehic.*', s) or re.search('electric.* motor.*', s) 
        or re.search('hybrid motor.*', s) or re.search('hybrid driv.*', s) or re.search('electric.* car.*', s)
        or re.search('hybrid car.*', s) or re.search('electric.* machin.*', s) or re.search('electric.* auto.*', s) 
        or re.search('hybrid auto.*', s) or re.search('yaw.* rat.* sens.*', s)):
        
        score += 1
        print('elc - alt fuel veh hit')
    
    #alt fuels
    if (re.search('alternat.* fuel.*', s) or re.search('mainstream.* fuel.*', s) or re.search('fuel cell.*', s) or re.search('nuclear powe.*', s) 
        or re.search('nuclear stat.*', s) or re.search('nuclear plant.*', s) or re.search('nuclear energ.*', s) 
        or re.search('nuclear and electric.*', s) or re.search('nuclear fuel.*', s) or re.search('fuel.* process.*', s) 
        or re.search('porous.* struct.*', s) or re.search('porous.* substrat.*', s) or re.search('solid.* oxid.* fuel.*', s) 
        or re.search('Fischer.*Tropsch.*', s) or re.search('refus.* deriv.* fuel.*', s) or re.search('refus.*deriv.* fuel.*', s) 
        or (re.search('fuel.*', s) and re.search('biotech.*', s) and (re.search('ethanol.*', s) or re.search('hydrogen.*', s))) 
        or re.search('bio.*fuel.*', s) or re.search('synthetic fuel', s) or re.search('combined heat and power', s) 
        or re.search('synth.* gas.*', s) or re.search('syngas', s)):
        
        score += 1
        print('elc - alt fuel hit')
    
    #electrochem
    if (re.search('electrochem.* cell.*', s) or re.search('electrochem.* fuel.*', s) or re.search('membran.* electrod.*', s) 
        or re.search('ion.* exchang.* membran.*', s) or re.search('ion.*exchang.* membran.*', s) or re.search('electrolyt.* cell.*', s) 
        or re.search('catalyt.* convers.*', s) or re.search('solid.* separat.*', s) or re.search('membran.* separat.*', s) 
        or re.search('ion.* exchang.* resin.*', s) or re.search('ion.*exchang.* resin.*', s) or re.search('proton.* exchang.* membra.*', s) 
        or re.search('proton.*exchang.* membra.*', s) or re.search('cataly.* reduc.*', s) or re.search('electrod.* membran.*', s) 
        or re.search('therm.* engin.*', s)):
        
        score += 1
        print('elc - electrochem hit')
    
    #battery
    if ((re.search('batter.*', s) or re.search('accumul.*', s)) and (re.search('charg.*', s) or re.search('rechar.*', s) 
        or re.search('turbocharg.*', s) or re.search('high capacit.*', s) or re.search('rapid charg.*', s) or re.search('long life', s) 
        or re.search('ultra.*', s) or re.search('solar', s) or re.search('no lead', s) or re.search('no mercury', s) 
        or re.search('no cadmium', s) or re.search('lithium.*ion.*', s) or re.search('lithium.* ion.*', s) or re.search('Li.*ion', s))):
        
        score += 1
        print('elc - battery hit')
    
    #addl energy sources
    if (re.search('addition.* energ.* sourc.*', s) or re.search('addition.* sourc.* of ener.*', s)):
        score += 1
        print('elc - addl energy hit')
    
    #carbon capture
    if ((re.search('carbon', s) and re.search('captu.*', s)) or (re.search('carbon', s) and re.search('stor.*', s)) 
        or re.search('carbon dioxid.*', s) or re.search('CO2', s)):
        
        score += 1
        print('elc - carbon capture hit')
    
    #energy manage
    if (re.search('ener.* sav.*', s) or re.search('ener.* effic.*', s) or re.search('energ.*effic.*', s) or re.search('energ.*sav.*', s) 
        or re.search('light.* emit.* diod.*', s) or re.search("^\W*LED", s) or re.search('organic LED', s) or re.search('OLED', s) 
        or re.search('CFL', s) or re.search('compact fluorescent.*', s) or re.search('energ.* conserve.*', s)):
        
        score += 1
        print('elc - energy manage hit')
    
    #building tech
    if ((re.search('build.*', s) or re.search('construct.*', s)) and (re.search('insula.*', s) or re.search('heat.* retent.*', s) 
        or re.search('heat.* exchang.*', s) or re.search('heat.* pump.*', s) or re.search('therm.* exchang.*', s) 
        or re.search('therm.* decompos.*', s) or re.search('therm.* energ.*', s) or re.search('therm.* communic.*', s) 
        or re.search('thermoplast.*', s) or re.search('thermocoup.*', s) or re.search('heat.* recover.*', s))):
        
        score += 1
        print('elc - building tech hit')
    
    return score

#--Helper function for selenium config and start--
def start_selenium():
    #start selenium
    options = Options();
    options.preferences.update({"javascript.enabled": False,"browser.link.open_newwindow": 1, "browser.link.open_newwindow.restriction": 0, "permissions.default.image": 2, "extensions.pocket.enabled": False, "browser.display.show_image_placeholders": False, "browser.display.use_document_fonts": 0, "media.volume_scale": 0}) #disabling JS helps find links
    driver = webdriver.Firefox(options=options)
    driver.install_addon('/home/user/pycleantech/ublock.xpi', temporary=False)
    driver.set_page_load_timeout(10)
    driver.maximize_window()
    
    return driver

#--Helper function to check connectivity--
def check_connectivity():
    while(True):
        try:
            urllib.request.urlopen(CONNECTIVITY_CHECK_URL)
            return #if above line completes successfully, then return and end
        except: #if fails
            print('Connection failed, trying again in 30 seconds.')
            time.sleep(30)

#--Begin main program--

df = pd.read_excel(EXCEL_WORKBOOK_NAME, sheet_name=EXCEL_SHEET_NAME) #read the excel file with company urls
df1 = df.where(pd.notnull(df), None) #replaces all NaN with None

#create output workbook
wb = Workbook()
sheet1 = wb.add_sheet(OUTPUT_SHEET_NAME) #add sheet to output workbook

#start selenium
driver = start_selenium()

#counts what row we are on
row = 0

#iterate over addresses in rows
for ind in range(STARTING_ROW, max(df1.index) + 1): #for starting index, subtract 2 from row number in Excel file.
    #check connectivity!
    check_connectivity()

    #start score counter
    desScore = 0
    
    if df1[WEBSITE_ADDRESS_COLUMN][ind] != None: #makes sure that a url is actually in the cell
        url = "https://" + df1[WEBSITE_ADDRESS_COLUMN][ind] + '/' #formats url
        
        print(url) #prints url to console
        sheet1.write(row, 5, str(url)) #writes url to output sheet for easier cross-checking
        
        try:
            driver.get(url) #visit url

            urls = [] #creates storage array for found links
            
            elems = driver.find_elements_by_xpath("//a[@href]") #searches for links by html tag
            
            for elem in elems:
                #note: str(link.get('href')).count('/') <= url.count('/') + 2 makes sure that even if root url is domain.com/a, then can still go down 2 more levels
                link = elem.get_attribute("href")
                #below is very interesting. Apparently, there can be links with nothing in them, just blank, that will throw a out of index error on link[0] line, so we need to filter that.
                #also made the blocked wordlist check smaller using the array.
                if len(link) > 0 and link.count('/') <= url.count('/') + 2 and not any(word in link for word in ['career', 'job', 'privacy', 'contact', '?', '#', 'terms', 'file', 'login', 'tel:', 'tell:', 'mailto:', 'fax:', 'ftp', 'magnet:', 'maps:', 'sms:', 'pdf', '.zip', 'facebook', 'twitter', 'youtube', 'instagram', 'linkedin']):
                    if link[0] == '/': #meaning relative (ex. /about-me), so must convert to absolute
                        urls.append(url + link)
                    else: #meaning absolute link
                        urls.append(link)
            
            final_urls = list(set(urls)) #removes duplicates
            print(final_urls) #prints array to console
            
            if final_urls == []: #means no subdirs found
                desScore = -2
            if len(final_urls) > 15: #meaning more than 15 urls
                final_urls = final_urls[0:14] #truncates down to 15 urls
            
            for path in final_urls:
                try:
                    #get page text using selenium and pyautogui select all
                    driver.get(path)
                    #below has been changed. In order to avoid activating links that may be under mouse cursor, right click is pressed instead, then escape key.
                    pyautogui.click(x=1, y=200, button = 'right')
                    pyautogui.press('esc')
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.hotkey('ctrl', 'c') #copy text
                    pagetext=pyperclip.paste() #paste text into variable

                    #below has been altered to prevent out of index errors.
                    if len(pagetext) > 10:
                        print(pagetext[0:9]) #prints first 10 characters to console
                    else:
                        print(pagetext) #prints entire text to console, if 10 or less than 10 chars long 
                    
                    #add values returned from helper functions to score variable
                    desScore = desScore + general(pagetext)
                    desScore = desScore + environ(pagetext)
                    desScore = desScore + renew(pagetext)
                    desScore = desScore + elc(pagetext)

                except Exception as e:
                    sheet1.write(row, 2, 'One or more subpages failed to be fetched. Error: ' + str(e))
                        
        except Exception as e: #meaning something fails
            desScore = -1 #flags for later manual intervention
            sheet1.write(row, 1, str(e)) #writes exception for easy debug
            print(e)
   
    if df1['Website address'][ind] == None: #meaning no url
        desScore = -1 #flags
        sheet1.write(row, 1, 'No URL.')

    print(desScore) #prints score to console
    sheet1.write(row, 0, desScore) #writes score to sheet
    
    row = row + 1 #increment row by one
    wb.save(OUTPUT_WORKBOOK_NAME)
    
    #below is unnecessary if doing on high-memory machine, this is only for low-memory systems.
    if row % 50 == 0: #meaning row is multiple of 50, close and re-start selenium to curb memory leak
        driver.quit()
        driver = start_selenium()
        
wb.save(OUTPUT_WORKBOOK_NAME) #save workbook
driver.quit() #quit selenium
