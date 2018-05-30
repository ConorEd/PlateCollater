import sys, os, time
import csv
import xlrd
import numpy as np
from tkinter import Tk
from tkinter import filedialog
import glob
import codecs
import xlsxwriter as xl
from os.path import isfile, join
from stat import S_ISREG, ST_MTIME, ST_MODE

#Open Data Directory
Tk().withdraw()
dirpath = filedialog.askdirectory()                                                       #Query the location of .asc files

#Maximum Plate Count and Minimum Time Interval
MaxplateNum = 10                                                                          
readTime = 5                                                                                            

#Query Plate Number and Time Interval # 
plateNum = int(input('Enter number of plates to be collated (1-10): '))                                 
if plateNum > MaxplateNum: 
	raise ValueError('A valid number of plates is required (10 = MAX)!')
MinInterval = plateNum*readTime
MinStr =str(MinInterval)
MinPrint = str(" (Minimum = "+MinStr+ " min): ")
Timepoint = int(input('Enter time interval used in minutes' + MinPrint))
if Timepoint < MinInterval:
	raise ValueError('Specified interval is below defined minimum!')

#Cast Data to Array and Order Chronologically
fnames = os.listdir(dirpath)
fnames.sort(key=lambda x: os.stat(os.path.join(dirpath, x)).st_mtime)                     

#Print Date of First Experiment (User Check Essentially)
firstfile = fnames[0]                                                                     #The .asc filenames include the date
#print(firstfile)
day = firstfile[4:6]
month = firstfile[2:4]
year = firstfile[0:2]
dateString = ('This Experiment begun on ' + day + '/' + month + '/20' + year)
print(dateString)

#Make Output Directory
savedir = dirpath+"/"+day+month+year+"-O"                                                 #Use the date as identifier for output directory
if not os.path.exists(savedir):
    os.makedirs(savedir)

#Data Collation Variables
plateRange = list(range(0, plateNum))                                                     #Number of plates to divide by
ascCount = len(glob.glob1(dirpath, "*.asc"))                                              #Determine number of '.asc' files
print(str(ascCount) + " ascii files found!" )
print("Collating...")
ascpPlate = (ascCount//plateNum)                                                          #Determine number of timepoints ('.asc'/plate
plateLabel = 0
d={}                                                                                      #Initialise dictionary
PlateCol = ["A", "B", "C", "D", "E", "F", "G", "H"]                                       #Used to map the 96 well plate into .xlsx
Wlab = 1

#Organise Files into a Dictionary based on Plate Number
for x in range(0, plateNum):
    d["Plate{0}".format(x)]=fnames[plateRange[plateLabel]::plateNum]                      
    plateLabel = plateLabel +1

#Set the Collation Loop Variables
plateName = 1
plateCount = 0
timCol = 3
timeInt = 0
colxl = 1
ascIter = 0 
p="Plate"
#w = 0 
rowNum = 2
colNum = 1
rawvar = 1
raw = []                                                                                  #Initialise array that will contain current plate's data.

#The Collation Loop
for x in range(0, plateNum):
    p1 = (p + str(plateCount))                                                              
    p2 = (p + str(plateName))                                                             #As dictionary keys are 0 indexed, we'll need a seperate variable for naming plate output files
    workbook = xl.Workbook(savedir + "/" + p2 + '.xlsx')
    worksheet1 = workbook.add_worksheet()
    worksheet1.write('A1', p2)
    worksheet1.write('A2', 'Timepoints')
    worksheet1.write('B1', 'WellID')
    platex = [*d[p1]]                                                                     #Convert the current plate from dictionary key to list object
    for x in range(0, 12):         
        PlateMap = [s + str(Wlab) for s in PlateCol] 
        worksheet1.write_row(1, colxl, PlateMap)
        Wlab+=1	                                                                              
        colxl+=8
    for x in range(0, ascpPlate):
        worksheet1.write(("A" + str(timCol)), str(timeInt))
        timCol+=1
        timeInt = timeInt + Timepoint
        with codecs.open(dirpath + '/' + platex[ascIter], encoding='utf-8-sig') as f:     #Open '.asc' Files
            raw=[[str(x) for x in line.split()] for line in f]
            for x in range(0, 96):                                                        #Loop that does one timepoint
                rawset = raw[rawvar]
                rawread = rawset[1]
                worksheet1.write(rowNum, colNum, rawread)                                 #Write one cell to output .xlsx
                colNum+=1
                rawvar+=1
            colNum=1
            rowNum+=1
        rawvar=1
        ascIter+=1
        
#Close the File for Plate(n)
    plateCount += 1                                                                       #Shift loop to next plate
    plateName += 1
    workbook.close()    
    print(p2 + ' = Done!')                                                              
    
#Reset Collation Loop Variables
    timCol = 3
    timeInt = 0
    Wlab = 1
    colxl = 1
    ascIter = 0
    rowNum = 2
    colNum = 1 
print("Operation Finished! Saved alongside your .asc files in Folder " + day + month + year + "-O. If there are any issues with the script let me know.")
