# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 11:44:07 2020

@author: mishaun
"""

import os, re, openpyxl, calendar
import pandas as pd

#function to remove old .xls files after running macro to convert xls to xlsx
def cleanup(files):
    for name in files:
        #checks multiple places after splitting at . for files that have . at the end of name (for ex: blackstone files)
        #try and except statement to prevent break on index out of bounds 
        try:
            if name.split('.')[1] == 'xls' or name.split('.')[2] == 'xls' or name.split('.')[3] == 'xls': 
                os.remove(gauge_sheet_directory + "\\" + name)
            else:
                continue
        except:
            continue

def getSum(cells):
    '''Parameter cells must be an openpyxl tuple of cell object
       This function will return the sum of cells passed in 
    '''
    temp = []
    
    #looping through tuple of cells, accessing first tuple, and then cell object within tuple
    for x in range(0,len(cells)):
        #prevents error of summing list if there is no value or nonetype in range of cells
        if type(cells[x][0].value)!= type(None): 
            temp.append(cells[x][0].value)
    
    return round(sum(temp))

def getStockSum(stocknumbers):
    
    '''
    This function will sum stock of cells and round
    stocknumbers parameter will be openpyxl cell addresses 
    sheet is global variable and an openpyxl object
    '''
    
    stock = 0
    for cell in stocknumbers:
        stock += sheet[cell].value
    return round(stock)
 
def special_remove(workbook, sheetindex):
    '''
    This function will remove a sheet in order for the correct sheet to be selected in selection portion of program
    '''
    ss_sheet = workbook[workbook.sheetnames[sheetindex]]
    #renaming sheet based on first instance before there is a comma, space, etc
    workbook.remove(ss_sheet)
    workbook.save(gauge_sheet_directory + '\\' + gauge_sheet_filename)
    
    
    
#Parameters for input based on report month
reportMonth = "Nov"
reportYear = 2019
monthDict={'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}


#Determining days in month - Used for shifting values of cells if needed
days = calendar.monthrange(reportYear, monthDict[reportMonth])[1]


#Path for gauge sheet folder
gauge_sheet_directory = 'F:\MAGNUM\Mishaun\Python Scripts\Data Extraction\Gauge Sheets - November 2019'

#Storing file names for gauge sheets in variable filenames
filenames = os.listdir(gauge_sheet_directory)

#importing data map dataframe
data_map_path = 'F:\MAGNUM\Mishaun\Production Reporting\Mishaun Notes'
data_map = pd.read_excel(data_map_path + '\\' + 'Data Map - Gauge Sheets - MPLP.xlsx', header = 6, usecols = 'C:AE')


#limiting data map to the wells that have inputted info based on oil production last month
map_filtered = data_map[(data_map['FilterList']==1) & (data_map["Filename"] != 'Not Available')]

for x in range(0,len(map_filtered)):
    #slicing off 1 well and assigning values needed for determining which functions to call
    well = map_filtered.iloc[x]
    
    ######### Excel opening, retrieving, closing section
    
    #opening workbook - dataonly flag set to true to get values only from cells
    
    gauge_sheet_filename = well["Filename"]
    wb = openpyxl.load_workbook(gauge_sheet_directory + '\\' + gauge_sheet_filename, data_only = True)
    
    #special case to cleanup vieman gauge sheet names for data to be found correctly - removing second sheet in gaugesheet which is useless
    if well["RRC ID"] == 21610:
        try:
            special_remove(wb,1)
        except:
            pass
    
    
    sheetnames = wb.sheetnames
    
    
    
    
    ######Picking Correct Sheet in Workbook Section
    
    #variable will be assigned in one of the if statements
    getsheet = ""
    
    #checking condition if well has gauge sheets by month and year
    if well["Monthly Tabs"]=="Y" and well["Multiple Years"] == "Y":
        
        for name in sheetnames:
            if reportMonth.upper() in name.upper() and str(reportYear)[-2:] in name:
                getsheet = name
                break
        
        #This will check the workbook for tab labeled with report month if the above statement fails
        if len(getsheet) == 0:
            for name in sheetnames:
                if reportMonth.upper() in name.upper():
                    getsheet = name
                    break 
            
    #checking condition if well has gauge sheets by month
    elif well["Monthly Tabs"]=="Y" and well["Multiple Years"] == "N":
        
        for name in sheetnames:
            if reportMonth.upper() in name.upper():
                getsheet = name
                break
    
    #Checking condition if well does not have monthly tabs, then look for well name
    elif well["Monthly Tabs"] == "N":
        
        if len(sheetnames) == 1:
            getsheet = sheetnames[0]
        else:
            #taking well name and splitting at nonalphanumeric characters
            regex = re.split('\W', well["Lease Name"].upper())
            
            matches = []
            #taking each instance of the split well name and searching through each sheetname
            for item in regex:
                for name in sheetnames:
                    #Only looking at first 75% of letters of regex
                    item_length = round(len(item)*0.75)
                    match = re.findall(item[0:item_length], name.upper())
                    #if the wellname regex is found, append the name of the sheet to the matches array
                    if match:
                        matches.append(name)
                #if matches were found/appended, then we can break the regex search loop 
                if len(matches)>0:
                    break
            #checking to see if there were multiple matches, if 1 match, then get the only sheet
            if len(matches) < 2:
                getsheet = matches[0]
            
            #if there are more than 1 match, pull the well number using regex library and search the matches list for that number
            elif len(matches)>=2:
                for i in range(0,len(matches)):
                    #if the string value of well number = digit found in sheetname, break loop and assign sheetname to getsheet
                    if str(well["Well Number"]) == re.findall('\d', matches[i])[0]:
                        break
                
                getsheet = matches[i]
                    
            
    
    #assigning found sheet to sheet            
    sheet = wb[getsheet]
    
    #####End Picking Sheet Section
    
    ####Getting Closing Stock Section
    
    #splitting data in stock cell at comma to account for wells that require adding multiple tanks
    stockCells = well["Oil Stock Cell"].split(',')
    closing_stock=0
    
    #if no shift is required, get values in stock cells and sum and round    
    if well["Shift Up Required"] == "N":
        closing_stock = getStockSum(stockCells)
    
    else:
        shiftedStockCells = []  
        
        for cell in stockCells:
            #re.findall returns list, so indexing string value at 0
            cellNumber = re.findall(r'\d+', cell)[0]
            
            #shifting value of cell up by however many days there in the month against default 31 days and converting back to string
            cellNumber = str(int(cellNumber) - (31-days))
            
            #extracting letter from oil stock cells to concat with shifted cell number
            cellLetter = re.findall('[A-Z]',cell)[0]
            
            shiftedStockCells.append(cellLetter + cellNumber)
        
        closing_stock = getStockSum(shiftedStockCells)
        
        
            
    ##############Pulling gauge sheet production section       
    
    #pulling gauge sheet production by passing in range of cells needed to be summed from datamap
    start_cell = well["Gauge Sheet Oil Production"].split(",")[0]
    end_cell = well["Gauge Sheet Oil Production"].split(",")[1]
    gauge_prod_cells = sheet[start_cell:end_cell]
    gauge_prod = round(getSum(gauge_prod_cells))
    
    #passing in value to main data frame using loc to avoid SettingWithCopyWarning
    data_map.loc[data_map["RRC ID"] == well["RRC ID"], "Closing Oil Stock"] = closing_stock
    data_map.loc[data_map["RRC ID"] == well["RRC ID"], "Gauge Sheet Prod Vol"] = gauge_prod


#adding two columns to check manual vs automated extraction differences
data_map["Gauge Prod Diff"] = data_map["Gauge Sheet Prod Vol"] - data_map["Manual Extraction Gauge Sheet Prod"]
data_map["Closing Stock Diff"] = data_map["Closing Oil Stock"] - data_map["Manual Extraction Closing Stock"]

#updating filtered version of full data_map
map_filtered = data_map[(data_map['FilterList']==1) & (data_map["Filename"] != 'Not Available')]

checker = map_filtered[["Lease Name", "Gauge Sheet Prod Vol","Manual Extraction Gauge Sheet Prod", "Closing Oil Stock","Manual Extraction Closing Stock","Gauge Prod Diff", "Closing Stock Diff"]]