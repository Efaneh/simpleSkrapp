import streamlit as st 
import csv
from PIL import Image
import time
import io
import sharepy
import pandas as pd
import openpyxl


trimCount = [0,0,0,0]
tabLogo = Image.open('tabLogo.png')
realLogo = Image.open('realLogo.png')
stonks = Image.open('stonks.png')
filterSizes = True
filterIndustries = True
bandsToUse = []
       
def getTopline(csvFile):
    topLine = csvFile.readline()
    topLine = topLine.replace("\n","")
    topLine = topLine.replace("\r","") #Removes all carrige return/newlines to make ensure CSV formats correctly
    return topLine    

def createLists(csvFile):
    topLine = getTopline(csvFile)
    topLineList = topLine.split(",") #get keys for dict
    topLineList = renameHeaders(topLineList)
    dictList = []
    csvReader = csv.reader(csvFile, quotechar='"', delimiter=',')
    for line in csvReader:
        dictList.append(dict(zip(topLineList, line))) # For each line in the csv file - create a dictionary
    dictList.pop(0) #Remove topline to avoid 'First name' : 'First name' etc.
    return dictList
    
def renameHeaders(headerList):
    index = headerList.index('Company website') if 'Company website' in headerList else -1
    if index != -1:
        headerList[index] = 'Company URL'
    index = headerList.index('Company Industry') if 'Company Industry' in headerList else -1
    if index != -1:
        headerList[index] = 'Industry'  
    return headerList
    
    
    
def cleanFirstName(dictList):
    for i in range(0,len(dictList)):
        firstName = dictList[i].get('First name').replace('-',' ')              #Replaces - with ' '
        firstName = firstName.replace('.',' ') 
        firstName = firstName.replace(',',' ')
        firstName = firstName.title()                       #Ensures first letter is capitalised, none others
        firstName = firstName.strip()                       #Removes leading/trailing whitespace                    
        splitName = firstName.split(' ')                    #Splits on whitespace, to remove double barrel etc.
        if splitName[0] == 'Dr':                            #If Dr X, take the X, else take first word in first name
            firstName = splitName[1]
        else:
            firstName = splitName[0]
        dictList[i].update({'First name': firstName}) 
    return dictList
       
       
def cleanLastName(dictList, removeNames):
    length = len(dictList)
    i = 0
    count = 0
    while i < length:
        lastName = dictList[i].get('Last name').replace(',',' ') 
        lastName = dictList[i].get('Last name').replace('.',' ') 
        lastName = lastName.title()
        lastName = lastName.strip()
        splitName = lastName.split(' ')
        if len(splitName[0]) == 1 and len(splitName) > 1: #If J. Smith, take the Smith
            lastName = splitName[1]
            dictList[i].update({'Last name': lastName})
        elif len(splitName[0]) == 1 and len(splitName) == 1 and removeNames: #If J. and remove tag - delete this person
            dictList.remove(dictList[i])
            i -= 1
            length -= 1
            count += 1
        elif (str.isspace(lastName) or lastName == '' or lastName is None) and removeNames: #Could create some complicated condition for the sake of not repeating code
            dictList.remove(dictList[i])                                                    #Removes empty surnames
            i -= 1
            length -= 1
            count += 1   
        else:
            lastName = splitName[0]  #If falls through to here, update the name by taking the first word of the surname
            dictList[i].update({'Last name': lastName})
        i += 1    
    trimCount[3] = count  #For stats later when counting # of deleted rows
    return dictList   #Potential issue in this code if their surname is 'Harry Jones', and they have not hyphenated, it will take Harry 
                      #Could always take the last word? (After checking it's not single digit of course) - Ask Dom for input
    
    
def checkForRepeats(dictList): #Alerts the SDR if there are more than 5 leads from the same account
    listOfCompanies = []
    fiveOrMore = []
    for i in range(0,len(dictList)):
        listOfCompanies.append(dictList[i].get('Company')) #Create list of comapnies.
    setOfCompanies = set(listOfCompanies)    
    for companyName in setOfCompanies:
        if listOfCompanies.count(companyName) > 5:
            fiveOrMore.append(companyName) #If company appears five times or more POST cleaning, alert SDR.
    return fiveOrMore        
    

def cleanDictList(dictList, streamInfo):
    excludedKeys = ['Company Founded','Company Headquarters','Email Status','Location']
    dictList = splitLocation(dictList)
    for i in range(0,len(dictList)):
        for j in range (0, len(excludedKeys)):
            if excludedKeys[j] in dictList[i].keys(): #If 'Company Founded' for example, is in the csv, delete that row
                dictList[i].pop(excludedKeys[j])
    if streamInfo[1]:        
        dictList = trimCompanySizes(dictList) #Trim company sizes based on banding input (Could change variable name)  
    if streamInfo[0]:
        dictList = trimIndustries(dictList, streamInfo[2])         #Trim company sizes based on SDR name assigned industries
    dictList = trimOppsAndCustomers(dictList)            #Trim Open opps and customers based on custom SalesForce report giving domain of open opps/customers 
    return dictList                                          #Consider a way to export this csv automatically once per day to github repo.

def splitLocation(dictList):
    cityListUK = populateList('cityListUK')
    cityListIreland = populateList('cityListIreland')
    stateListUK = populateList('stateListUK')
    countryList = populateList('countryList')
    for i in range(0,len(dictList)):
        dictList[i]['City'] = ''
        dictList[i]['State'] = ''
        dictList[i]['Country'] = ''
        location = dictList[i].get('Location', 'Deleted')
        if location == 'Deleted':
            continue
        for country in countryList:
            if country in location:
                dictList[i]['Country'] = country 
                break
        for state in stateListUK:
            if state in location:
                dictList[i]['State'] = state
                break
        for city in cityListUK:
            if city in location:
                dictList[i]['City'] = city
                break           
        if dictList[i].get('City') != '':
            dictList[i]['Country'] = 'United Kingdom'
        for city in cityListIreland:
            if city in location:
                dictList[i]['City'] = city
                break       
    if location == 'Deleted':
        st.error('Note that the location column had been deleted - so all location data will be left blank')
    return dictList
    
def populateList(string):
    with open(string+'.csv', 'r', encoding='utf-8') as file:
        allText = file.read()
        listToReturn = allText.split('\n')
    return listToReturn   

def trimOppsAndCustomers(dictList):
    count = 0
    j = 0
    s = sharepy.connect('kalliduslimited.sharepoint.com', 'efan.haynes@kallidus.com','clxnqltcptcvkhvg')
    string = s.get('https://kalliduslimited.sharepoint.com/sites/Sales/Shared%20Documents/Apps/currentCustomers.xlsx')
    f = io.BytesIO(string.content)
    dataFrame = pd.read_excel(f)
    urlList = dataFrame.values.tolist() 
    customerWebsiteList = []
    for url in urlList:
        customerWebsiteList.append(url[1])  #Create list of domains to match
    length = len(dictList)
    while j < length:
        if dictList[j].get('Company URL') in customerWebsiteList: #If there's a match, remove
            dictList.remove(dictList[j])
            count += 1
            j -= 1
            length -= 1
        j += 1        
    trimCount[2] = count #Stat tracking
    return dictList
    
      

def trimCompanySizes(dictList):
    i = 0
    count = 0
    length = len(dictList)
    checkForBanding(dictList)
    while i < length:
        if dictList[i].get('Company Size') not in bandsToUse:
            dictList.remove(dictList[i])
            i -= 1
            length -= 1
            count += 1
        i += 1
    trimCount[0] = count 
    return dictList
    

def checkForBanding(dictList):
    if 'Company Size' not in dictList[0].keys():
        st.error("It looks like you've deleted the company size column, please untick the \"Filter by company banding\" option!")
        st.stop()
    return        
    
    
def trimIndustries(dictList, sdrName):
    masterAccList = []
    name = False
    count = 0
    with open('sdrAccounts.csv', 'r', encoding='utf-8') as sdrFile:
        sdrReader = csv.reader(sdrFile, quotechar='"', delimiter=',')
        for line in sdrReader:
            if sdrName in line:
                masterAccList = line
                name = True           
    if name == True:
        i = 0
        length = len(dictList)
        while i < length:
            if dictList[i].get('Industry') not in masterAccList: #If company industry is not in the assigned industries, remove it
                dictList.remove(dictList[i])                             #Could look to send these to the correct SDR each week
                count += 1
                i -= 1
                length -= 1
            i += 1
    trimCount[1] = count       
    return dictList            
                
                    
def createSimpleSkrapp(dictList, fileName, name):
    try:
        nameList = name.split(' ')
        firstName = nameList[0].lower()
        lastName = nameList[1].lower()
        stringToWrite = ''
        keyList = dictList[0].keys()
        for key in keyList:
            stringToWrite += key + ',' #The topline
        stringToWrite += 'Owner\n'
        for i in range(0,len(dictList)):
            count = 0
            valuesInDict = dictList[i].values()
            for value in valuesInDict:
                count =  count+1
                stringToWrite += '"' + value + '",' #Append to this huge string that will become the new file with each column value
            stringToWrite += firstName + '.' + lastName + '@kallidus.com\n'
    except:
        st.error('It looks like simpleSkrapp has removed EVERY row - have you selected the correct name/banding?')
        st.stop()
    st.download_button("⬇️ Download File ⬇️",stringToWrite,fileName[:-4] + '_simpleSkrapped.csv')
    return fileName[:-4] + '_simpleSkrapped.csv'

def skrappReport(fileName, fiveOrMore):
    chartData = pd.DataFrame({
                    'index': ['Banding', 'Industry', 'Current Op/Customer', 'Surname'],
                    'Deleted Rows' : [trimCount[0],trimCount[1],trimCount[2], trimCount[3]], #Creating the graph
        }).set_index('index')            
    with st.expander('simpleSkrapp Report'):
        st.write('The tool is now finished! A new file ' + fileName + ' has been created for you and can be downloaded above.')
        st.write("Note that the deletion is done in this order, so if a row has the wrong banding AND industry it will only count towards banding for the sake of stats.")
        st.write("Rows deleted due to incorrect company banding - " + str(trimCount[0]))
        st.write("Rows deleted due to current Op/Customer - " + str(trimCount[2]))
        st.write("Rows deleted due to single-letter surname or no surname - " + str(trimCount[3]))
        st.write("Don't forget you still have to clean the job titles and double check the file! 😉")
        st.bar_chart(chartData)
        st.write('The following companies have 5 or more leads, take out the least appropirate people!') #Pulling all the stats and data we have in our array and displaying via streamlit
        if len(fiveOrMore) > 0:
            for company in fiveOrMore:
                st.markdown('- ' + company)
        else:
            st.write('You lucky 🦆! There are none!') 
    
def createNameList():
    sdrNames = []
    with open('sdrAccounts.csv', 'r', encoding='utf-8') as sdrFile:
        sdrReader = csv.reader(sdrFile, quotechar='"', delimiter=',')
        for line in sdrReader:
            sdrNames.append(line[0]) #Create possible selection of names from sdrAccounts file
    return sorted(sdrNames)
    
def sliderChange(sliderOptions, sizeSliderMin, sizeSliderMax):   
    selectedBands = []
    filteredBands = ['1-10','11-50','51-200','201-500','501-1000','1001-5000','5001-10000','10001+'] 
    st.write('You are currently keeping the following bands:')
    for i in range(0,len(sliderOptions)):
        if int(sliderOptions[i].strip('+')) >= int(sizeSliderMin.strip('+')) and int(sliderOptions[i].strip('+')) <= int(sizeSliderMax.strip('+')):
            selectedBands.append(i)
    selectedBands.pop()        
    for band in selectedBands:
        bandsToUse.append(filteredBands[band])
        st.write(filteredBands[band])     
    
def streamlitSetup(sdrNames): #Setting up the frontend with streamlit
    sliderOptions = ['1', '10', '50','200','500','1000','5000','10000','10001+']
    st.set_page_config(page_title = 'simpleSkrapp', page_icon = tabLogo)
    st.title('simpleSkrapp') 
    boolList = setupSidebar()
    with st.form('parameters'):
        st.write("Welcome to simpleSkrapp 😊! This tool was made by Efan Haynes to automate some aspects of the "
        "Skrapp process.\n\nIt has not been tested for every possibility " 
        "So please do manually check the csv file after executing!\n\nGiven the LinkedIn changes to industries, and the fact we now SalesNav Search our"
        " account lists, the filter by industry option has been removed, however please still select your name to ensure the correct SalesLoft owner")
        simpleSkrappExplained()
        name = st.selectbox('Please select your name',sdrNames)
        sizeSliderMin, sizeSliderMax = st.select_slider('Please Select the range of company sizes to keep', sliderOptions, value = ('50', '200'))
        submittedBanding = st.form_submit_button('🔄 Update Banding & Name 🔄',on_click = sliderChange(sliderOptions, sizeSliderMin, sizeSliderMax))      
    if sizeSliderMin == sizeSliderMax:
        st.error('Please do not pick the same value for the minimum and maxixmum value!')
        st.stop()
    with st.form('button'): 
        uploadedFile = st.file_uploader("Choose a file to simpleSkrapp", 'csv')
        submitted = st.form_submit_button("✨ Simplify! ✨")         
    return [boolList[0], boolList[1], name, sizeSliderMin, sizeSliderMax, uploadedFile, submitted, boolList[2]]
    
def setupSidebar(): #Setting up the frontend with streamlit, top bit removes "Made by streamlit"
    hide_menu_style = """
                      <style>
                      footer {visibility: hidden;} 
                      </style>
                      """
    with st.sidebar:
        st.image(realLogo, caption='An extremely original logo')
        st.header('simpleSkrapp Options')
        filterIndustries = False
        filterSizes = st.checkbox('Filter by company banding', False, disabled = True)
        filterNames = st.checkbox('Remove leads with single letter or no surname', True)
        devOptions = st.checkbox('devMode', False)
        if not devOptions: #Gives option to go into streamlit options (Mainly for me to play around with when testing)
            hide_menu_style ="""
                      <style>
                      #MainMenu {visibility: hidden;}
                      footer {visibility: hidden;}
                      </style>
                      """
        st.image(stonks, caption ='Your SQL after using simpleSkrapp')
        st.markdown(hide_menu_style, unsafe_allow_html=True)
    boolList = [filterIndustries,filterSizes, filterNames]    
    return boolList    
    
def simpleSkrappExplained(): #Text displayed what it does
    with st.expander('simpleSkrapp explained'):
        st.header('What is simpleSkrapp?')
        st.write('simpleSkrapp is a place where you can upload the csv files you get from skrapp, and have them cleaned for you!')
        st.header('What can it do?')
        st.markdown('So far, it currently:\n - Deletes and inserts columns to the csv as necessary\n - Cleans the prospects first and last name\n'
                    '- Removes any rows that are current customers\n - Removes any companies outside the selected banding\n - Takes the location information from Skrapp and enters it in a manner SalesLoft understands'
                    ' \n - Removes any prospects with a single digit surname \n - Lets you know if there are more than 5 entries from a single company')
        st.header('What can it not do?')            
        st.markdown('- Clean job titles! People like to call themselves all sorts of things, so I\'m afraid that\'s still your job!')
        st.markdown('- Clean names with 100% accuracy. Please look over the names and delete any further unsuitable prospects')
    
    
def convertFile(uploadedFile): #Making sure correct file encoding is used 
    csvFile = io.StringIO(uploadedFile.getvalue().decode("utf-8"))
    return csvFile

def streamlitLogic(streamInfo): #Controling the flow of the code based on options and submit button
    if streamInfo[6]: #Whether or not submit button clicked
        if streamInfo[5] and streamInfo[4] and streamInfo[3] and streamInfo[2]:
            with st.spinner('Simplifying your skrapp - please wait'):
                time.sleep(0.5)
                csvFile = convertFile(streamInfo[5])
                dictList = createLists(csvFile)
            return dictList
        else:
            st.error("Please select your name, banding, and upload a file!")     

def main(): #Main 
    sdrNames = createNameList()
    streamInfo = streamlitSetup(sdrNames)
    dictList = streamlitLogic(streamInfo)
    if streamInfo[5] and streamInfo[6]:
        try:
            fileName = streamInfo[5].name
            dictList = cleanFirstName(dictList)
            dictList = cleanLastName(dictList, streamInfo[7])
            dictList = cleanDictList(dictList, streamInfo)
            fiveOrMore = checkForRepeats(dictList)
            trueFileName = createSimpleSkrapp(dictList, fileName, streamInfo[2])
        except:
            st.error('There seems to be an issue simplifying your Skrapp - feel free to try again but contact Efan if this persists')
            st.stop()
        skrappReport(trueFileName, fiveOrMore)    
        st.success('Finished!')
    
if __name__ == '__main__':
    main()
