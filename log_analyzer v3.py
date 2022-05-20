# -*- coding: utf-8 -*-
"""
Created on Wed Apr 27 16:02:17 2022

@author: nrehman
"""

import pandas as pd
import xlsxwriter
from datetime import datetime
from datetime import timedelta
import numpy as np
import time
import os
from configparser import ConfigParser


##--------------READ INI----------------
thisfolder = os.path.dirname(os.path.abspath(__file__))
ini_path = os.path.join(thisfolder,'setup.ini')

parser = ConfigParser()
parser.read(ini_path)

# Create a dictionary of the variables stored under the "files" section of the .ini
files = {param[0]: param[1] for param in parser.items('files')}
fields = {param[0]: param[1] for param in parser.items('fields')}
downtime = {param[0]: param[1] for param in parser.items('downtime')}

print(files)
print(fields)

##--------------READ IN DATA----------------

filepath = files['input_path']

print('Analyzing log files, please wait...')

#Extract Columns
log = pd.read_csv(filepath)

# convert columns names to lowercase
#log.columns = list(map(str.upper,log.columns))


##--------------PRODUCTION TIME ----------------
#Combine Date and Time columns to creat standard date/time format
starttime = datetime.strptime(log[fields['date']][0] + ' ' + log[fields['time']][0], '%d/%m/%Y %I:%M:%S %p')
endtime = datetime.strptime(log[fields['date']][log.shape[0] - 1] + ' ' + log[fields['time']][log.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
elapsed = endtime - starttime

completedate = max(log[fields['time']])
print('Complete Date: ' + str(completedate))


##--------------YIELD/DEFECTS-----------------

runqty = log.shape[0]
print('Total Run Qty: '+ str(runqty))

produced = log[log[fields['tagged']] != 'TAGGED'].shape[0] 
print('Total Produced: ' + str(produced))

tagged = log[log[fields['tagged']] == 'TAGGED'].shape[0]
print('Total TAG: ' + str(tagged))

reproduced = log[log[fields['reproduced']] == 'RE-PRODUCED'].shape[0]
defectrate = (tagged/runqty)*100

firstpass = log[log[fields['reproduced']] != 'RE-PRODUCED'] #total run qty in first pass
firstgood = firstpass[firstpass[fields['tagged']] != 'TAGGED'].shape[0]
print('First Pass Yield: ' + str(firstgood))

firstbad = firstpass[firstpass[fields['tagged']] == 'TAGGED'].shape[0]
print('First Pass Defects: ' + str(firstbad))

secondpass = log[log[fields['reproduced']] == 'RE-PRODUCED']
secondgood = secondpass[secondpass[fields['tagged']] != 'TAGGED'].shape[0]
print('Second Pass Yield: ' + str(secondgood))

secondbad = secondpass[secondpass[fields['tagged']] == 'TAGGED'].shape[0]
print('Second Pass Defects: ' + str(secondbad))

#create new df with count of all sleeve tag reasons 
tags = log[fields['tag_reason']].value_counts().reset_index(inplace=False)
tags.columns = ['Tag_Reason', 'Count']

#create list of percent TAG (each tag reason / total TAG)
defectpercent = []
if tagged > 0: #prevents divsion by 0
    for i in tags.Count:
        defectpercent.append((i/tagged)*100)


#add list of defect percentages as a column to df
tags['percent'] = defectpercent 
print(tags)


##--------------CREATE EXCEL WORKBOOK----------------
output = files['output_path'] + "/" + files['output_filename'] + ".xlsx"
workbook  = xlsxwriter.Workbook(output)
yieldsheet = workbook.add_worksheet('Yield')
downtimesheet= workbook.add_worksheet('Downtime')


##--------------DOWNTIME OUTPUT TO EXCEL---------------------
downtimesheet.write(0,0,'From')
downtimesheet.write(0,1,'To')
downtimesheet.write(0,2,'Downtime')
downtimerow = 1

total_downtime = timedelta(seconds=0)

for i in range(log[fields['time']].shape[0]):
    threshold = timedelta(seconds=int(downtime['threshold']))
    if i >0:
        time1 = datetime.strptime(log[fields['time']][i-1], '%I:%M:%S %p')
        time2 = datetime.strptime(log[fields['time']][i], '%I:%M:%S %p')
        diff = time2-time1

        if diff > threshold:            
            downtimesheet.write(downtimerow,0,str(time1))
            downtimesheet.write(downtimerow,1,str(time2))
            downtimesheet.write(downtimerow,2,str(diff))
            total_downtime += diff
            downtimerow += 1            

downtimesheet.write(downtimerow,1,'TOTAL: ')
downtimesheet.write(downtimerow,2,str(total_downtime))

perc_downtime = (total_downtime/elapsed)*100
perc_uptime = 100-perc_downtime

##-----------------YIELD OUTPUT TO EXCEL--------------------
desc = ['Start Date','Complete Time','Total Production Hrs','Total Downtime','% Downtime','% Uptime',
        'Total Run Qty','Total Produced','Total Tagged','Total Reproduced','Defect Rate (%)',
        'First Pass Yield','First Pass Defects','Second Pass Yield','Second Pass Defects']

data = [starttime,completedate,elapsed,total_downtime,perc_downtime,perc_uptime,runqty,produced,
        tagged,reproduced,defectrate,firstgood,firstbad,secondgood,secondbad]

for i in range(len(desc)):
    yieldsheet.write(i,0,desc[i])
    yieldsheet.write(i,1,str(data[i]))

yieldsheet.set_column(0,0,18)
yieldsheet.write(len(desc)+2,0,'Tag Reason',workbook.add_format({'bold': True}))
yieldsheet.write(len(desc)+2,1,'Count',workbook.add_format({'bold': True}))
yieldsheet.write(len(desc)+2,2,'%',workbook.add_format({'bold': True}))
    
for i in range(tags.shape[0]):
    yieldsheet.write(len(desc)+3+i,0,tags.Tag_Reason[i])
    yieldsheet.write(len(desc)+3+i,1,tags.Count[i])
    yieldsheet.write(len(desc)+3+i,2,tags.percent[i])



workbook.close()
os.startfile(output)
