import streamlit as st 
import pandas as pd
import numpy as np
import sys
import os
import csv
import re
from PIL import Image
import time
import io


trimCount = [0,0,0,0]
tabLogo = Image.open('tabLogo.png')
realLogo = Image.open('realLogo.png')
stonks = Image.open('stonks.png')
filterSizes = True
filterIndustries = True
       
def getTopline(csvFile):
    topLine = csvFile.readline()
    topLine = topLine.replace("\n","")
    topLine = topLine.replace("\r","")
    return topLine    

def createLists(csvFile):
    topLine = getTopline(csvFile)
    topLineList = topLine.split(",")
    dictList = []
    csvReader = csv.reader(csvFile, quotechar='"', delimiter=',')
    for line in csvReader:
        dictList.append(dict(zip(topLineList, line)))
    dictList.pop(0)
    return dictList
    
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
        if len(splitName[0]) == 1 and len(splitName) > 1:
            lastName = splitName[1]
            dictList[i].update({'Last name': lastName})
        elif len(splitName[0]) == 1 and len(splitName) == 1 and removeNames:
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
            lastName = splitName[0]       
            dictList[i].update({'Last name': lastName})
        i += 1    
    trimCount[3] = count        
    return dictList    
    
    
def checkForRepeats(dictList): #Alerts the SDR if there are more than 5 leads from the same account
    listOfCompanies = []
    fiveOrMore = []
    for i in range(0,len(dictList)):
        listOfCompanies.append(dictList[i].get('Company'))
    setOfCompanies = set(listOfCompanies)    
    for companyName in setOfCompanies:
        if listOfCompanies.count(companyName) > 5:
            fiveOrMore.append(companyName)
    return fiveOrMore        
    

def cleanDictList(dictList, streamInfo):
    excludedKeys = ['Company Founded','Company Headquarters','Email Status','Location']
    for i in range(0,len(dictList)):
        for j in range (0, len(excludedKeys)):
            if excludedKeys[j] in dictList[i].keys():
                dictList[i].pop(excludedKeys[j])
    if streamInfo[1]:        
        dictList = trimCompanySizes(dictList, streamInfo[3])    
    if streamInfo[0]:
        dictList = trimIndustries(dictList, streamInfo[2])    
        dictList = trimOppsAndCustomers(dictList)
    return dictList       

def trimOppsAndCustomers(dictList):
    count = 0
    j = 0
    with open('customerListClean.csv', 'r', encoding='utf-8') as file:
        customerListDict = createLists(file)
    customerWebsiteList = []
    for i in range (0, len(customerListDict)):
        customerWebsiteList.append(customerListDict[i].get("Domain"))  
    length = len(dictList)
    while j < length:
        if dictList[j].get('Company website') in customerWebsiteList:
            dictList.remove(dictList[j])
            count += 1
            j -= 1
            length -= 1
        j += 1        
    trimCount[2] = count
    return dictList
    
      

def trimCompanySizes(dictList, companySize):
    filteredBands = ['Self-employed','1-10','11-50','51-200','201-500','501-1000','1001-5000','5001-10000','100001+']
    i = 0
    count = 0
    length = len(dictList)
    if companySize == 'SME and below':
        filteredBands = ['1-10','11-50','51-200']
    elif companySize == 'SME':
        filteredBands = ['51-200']
    elif companySize == 'MM':
        filteredBands = ['201-500','501-1000','1001-5000']
    elif companySize == 'Enterprise':
        filteredBands = ['5001-10000','10001+']
    checkForBanding(dictList)
    while i < length:
        if dictList[i].get('Company Size') not in filteredBands:
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
            if dictList[i].get('Company Industry') not in masterAccList:
                dictList.remove(dictList[i])
                count += 1
                i -= 1
                length -= 1
            i += 1
    trimCount[1] = count       
    return dictList            
                
                    
def createSimpleSkrapp(dictList, fileName):
    try:
        stringToWrite = ''
        keyList = dictList[0].keys()
        for key in keyList:
            stringToWrite += key + ','
        stringToWrite += 'Contact Source\n'
        for i in range(0,len(dictList)):
            count = 0
            valuesInDict = dictList[i].values()
            for value in valuesInDict:
                count =  count+1
                stringToWrite += '"' + value + '",'
            stringToWrite += 'Hunter/Skrapp\n'
    except:
        st.error('It looks like simpleSkrapp has removed EVERY row - have you selected the correct name/banding?')
        st.stop()
    st.download_button("⬇️ Download File ⬇️",stringToWrite,fileName[:-4] + '_simpleSkrapped.csv')
    return fileName[:-4] + '_simpleSkrapped.csv'

def skrappReport(fileName, fiveOrMore):
    chartData = pd.DataFrame({
                    'index': ['Banding', 'Industry', 'Current Op/Customer', 'Surname'],
                    'Deleted Rows' : [trimCount[0],trimCount[1],trimCount[2], trimCount[3]],
        }).set_index('index')            
    with st.expander('simpleSkrapp Report'):
        st.write('The tool is now finished! A new file ' + fileName + ' has been created for you and can be downloaded above.')
        st.write("Note that the deletion is done in this order, so if a row has the wrong banding AND industry it will only count towards banding for the sake of stats.")
        st.write("Rows deleted due to incorrect company banding - " + str(trimCount[0]))
        st.write("Rows deleted due to incorrect industry - " + str(trimCount[1]))
        st.write("Rows deleted due to current Op/Customer - " + str(trimCount[2]))
        st.write("Rows deleted due to single-letter surname or no surname - " + str(trimCount[3]))
        st.write("Don't forget you still have to clean the job titles and double check the file! 😉")
        st.bar_chart(chartData)
        st.write('The following companies have 5 or more leads, take out the least appropirate people!')
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
            sdrNames.append(line[0])
    return sorted(sdrNames)
    
def streamlitSetup(sdrNames):
    st.set_page_config(page_title = 'simpleSkrapp', page_icon = tabLogo)
    st.title('simpleSkrapp') 
    boolList = setupSidebar()
    with st.form('parameters'):
        st.write("Welcome to simpleSkrapp 😊! This tool was made by Efan Haynes to automate some aspects of the "
        "Skrapp process.\n\nIt has not been tested for every possibility " 
        "So please do manually check the csv file after executing!")
        simpleSkrappExplained()
        name = st.selectbox('Please select your name',sdrNames)
        compSize = st.selectbox('Please enter the company banding:',['SME and below','SME','MM', 'Enterprise'])
        uploadedFile = st.file_uploader("Choose a file to simpleSkrapp", 'csv')
        submitted = st.form_submit_button("✨ Simplify! ✨")   
    return [boolList[0], boolList[1], name, compSize, uploadedFile, submitted, boolList[2]]
    
def setupSidebar():
    hide_menu_style = """
                      <style>
                      footer {visibility: hidden;}
                      </style>
                      """
    with st.sidebar:
        st.image(realLogo, caption='An extremely original logo')
        st.header('simpleSkrapp Options')
        filterIndustries = st.checkbox('Filter by industries', True)
        filterSizes = st.checkbox('Filter by company banding', True)
        filterNames = st.checkbox('Remove leads with single letter or no surname', True)
        devOptions = st.checkbox('devMode', False)
        if not devOptions:
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
    
def simpleSkrappExplained():
    with st.expander('simpleSkrapp explained'):
        st.header('What is simpleSkrapp?')
        st.write('simpleSkrapp is a place where you can upload the csv files you get from skrapp, and have them cleaned for you!')
        st.header('What can it do?')
        st.markdown('So far, it currently:\n - Deletes and inserts columns to the csv as necessary\n - Cleans the prospects first and last name\n'
                    '- Removes any rows that are outside of your allocated industries and selected banding\n - Removes any rows that are current customers'
                    ' \n - Removes any prospects with a single digit surname \n - Lets you know if there are more than 5 entries from a single company')
        st.header('What can it not do?')            
        st.markdown('- Clean job titles! People like to call themselves all sorts of things, so I\'m afraid that\'s still your job!')
        st.markdown('- Clean names with 100% accuracy. Please look over the names and delete any further unsuitable prospects')
    
    
def convertFile(uploadedFile):
    csvFile = io.StringIO(uploadedFile.getvalue().decode("utf-8"))
    return csvFile

def streamlitLogic(streamInfo):
    if streamInfo[5]: #Whether or not submit button clicked
        if streamInfo[4] and streamInfo[3] and streamInfo[2]:
            with st.spinner('Simplifying your skrapp - please wait'):
                time.sleep(0.5)
                csvFile = convertFile(streamInfo[4])
                dictList = createLists(csvFile)
            return dictList
        else:
            st.error("Please select your name, banding, and upload a file!")     

def main():
    sdrNames = createNameList()
    streamInfo = streamlitSetup(sdrNames)
    dictList = streamlitLogic(streamInfo)
    if streamInfo[5] and streamInfo[4]:
        try:
            fileName = streamInfo[4].name
            dictList = cleanFirstName(dictList)
            dictList = cleanLastName(dictList, streamInfo[6])
            dictList = cleanDictList(dictList, streamInfo)
            fiveOrMore = checkForRepeats(dictList)
            trueFileName = createSimpleSkrapp(dictList, fileName)
        except:
            st.error('There seems to be an issue simplifying your Skrapp - feel free to try again but contact Efan if this persists')
            st.stop()
        skrappReport(trueFileName, fiveOrMore)    
        st.success('Finished!')
    
if __name__ == '__main__':
    main()
