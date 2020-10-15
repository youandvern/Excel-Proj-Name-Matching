"""
Created on Tue Oct 13, 2020
@author: Andrew-V.Young

Match project desctiption to project number
"""


import sys
import pandas as pd
from openpyxl import load_workbook


# character matching count function 
# https://www.geeksforgeeks.org/python-count-the-number-of-matching-characters-in-a-pair-of-string/
def count(str1, str2):  
    c, j = 0, 0
    # loop executes till length of str1 and  
    # stores value of str1 character by character  
    # and stores in i at each iteration. 
    for i in str1:     
          
        # this will check if character extracted from 
        # str1 is present in str2 or not(str2.find(i) 
        # return -1 if not found otherwise return the  
        # starting occurrence index of that character 
        # in str2) and j == str1.find(i) is used to  
        # avoid the counting of the duplicate characters 
        # present in str1 found in str2 
        if str2.find(i)>= 0 and j == str1.find(i):  
            c += 1
        j += 1
    return  c 

# function to loop through pandas dataframe and find closest match between name provided and df description
def loop_rows(df, fullname, rmvword = ""):  # df=dataframe fullname = string of project name to match
    charcount = [] # initialize empty list to store # of matching characters
    
    # dataframe has columns: ProjNumber ProjName ProjType ProjCity ProjState
    # loop through rows in dataframe
    for index2, row2 in df.iterrows(): 
        # build df full name by combining name and city (remove first 6chars of name)
        namematch = str(row2[1])[6:] + " " + str(row2[3])
        # string project name to compare with df names (option to remove input substring)
        nummatch = count(fullname.replace(rmvword,""), namematch)
        charcount.append(nummatch)
    # locate position of closest match (last (most recent) instance if multiple)
    maxindex = len(charcount) - 1 - charcount[::-1].index(max(charcount)) 
    # compile output list (projNumber, dfName, inputName)
    list_append = [df.iloc[maxindex, 0], df.iloc[maxindex,1], fullname]
    return list_append


def reformat_number_table(inFileName = 'Book1.xlsx'):
    # function to reformat spreadsheet - matches sheet of job names to sheet of job numbers and descriptions
    
    # create dataframe for job number/description and job name sheets
    jobsTable = pd.read_excel(inFileName, sheet_name = 'Jobs', index_col = None, header = None).fillna(" ")
    nameTable = pd.read_excel(inFileName, sheet_name = 'Jname', index_col = None, header = None).fillna(" ")

    # initialize list of job numbers to export back to excel as final result
    numlist = []
    
    # extract separate dataframes for uniquely named locations
    guamJobs = jobsTable[jobsTable[4].str.contains("GUAM")]
    qatarJobs = jobsTable[jobsTable[4].str.contains("QATAR")]
    hiJobs = jobsTable[jobsTable[4].str.contains("HI")]
    saipanJobs = jobsTable[jobsTable[4].str.contains("SAIPAN")]
  

    # loop through each name from name df
    for index, row in nameTable.iterrows():
        
        # extract and format string name from df row
        fullname = row[0]
        names = row[0].split()
        name = names[0].strip(",()")
        
        # create new dataframe for job number/descriptions where the first word in the name is contained in a city name
        freshCol = jobsTable[3].str.contains(name).fillna(False)
        includedRows = jobsTable[freshCol]
        
        # count number of rows found
        numRows = len(includedRows.index)
        
        # initialize default job number result as Not Found
        list_append = ["Not Found","Not Found","Not Found"]
        
        if numRows == 1: # only one tank in the city searched for
            # append list (projName, projDescription, projName)
            list_append = [includedRows.iloc[0,0], includedRows.iloc[0,1], fullname]
         
        elif numRows == 0: # no city found in first word search --> unique location search
            if "Guam" in fullname or "GU" in fullname:
               list_append = loop_rows(guamJobs, fullname, rmvword="Guam")

            elif "Qatar" in fullname:
               list_append = loop_rows(qatarJobs, fullname)
               
            elif "Hawaii" in fullname or "HI" in fullname:
               # list_append = loop_rows(hiJobs, fullname)
               list_append = list_append
               
            elif "Saipan" in fullname:
               list_append = loop_rows(saipanJobs, fullname)
               
        elif numRows>1 and len(names) > 1:  # multiple results returned for a name with multple words
            
            # repeat search using first two words of input name
            testname = names[0].strip(",()") + " " + names[1].strip(",()")
            testfresh = jobsTable[3].str.contains(testname).fillna(False)
            testincluded = jobsTable[testfresh]
            testnumRow = len(testincluded.index)
            
            # if no results found from new search, find closest match from old search
            if testnumRow == 0: 
                list_append = loop_rows(includedRows, fullname)
                
            # if new search returns more than one result, find closest match from new search
            elif testnumRow > 1:
                # print(fullname + " --->>> " + testincluded)
                list_append = loop_rows(testincluded, fullname)
            
            # if new search returns single tank, use this project number
            elif testnumRow == 1:
                list_append = [includedRows.iloc[0,0], includedRows.iloc[0,1], fullname]

        # append result to list
        numlist.append(list_append)
        # if list_append == ["Not Found","Not Found","Not Found"]:
        #     print(fullname)
    
    # create dataframe from compiled results and print to excel
    numdf = pd.DataFrame(numlist, columns = ['Job Number', 'Job Name', 'Job Name 2'])  
     
    book = load_workbook(inFileName)  # new data entry without deleting existing

    # add sorted data to new sheet
    with pd.ExcelWriter(inFileName, engine = 'openpyxl') as writer:
        writer.book = book
        numdf.to_excel(writer, sheet_name = 'JNo')
        writer.save()
        writer.close()

    return 'reformatting complete'


# run function
run_funct = reformat_number_table('jobnames.xlsx')
print(run_funct)
