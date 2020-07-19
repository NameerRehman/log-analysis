# -*- coding: utf-8 -*-
import pandas as pd
import xlsxwriter
from datetime import datetime
import numpy as np
import time

#Read Files
sfile = input('Enter path of S log: ')
bfile = input('Enter path of B log: ')
cfile = input('Enter path of C log: ')
pfile = input('Enter path of P log: ')

print('\nAnalyzing log files, please wait...')

#Extract Columns
slog = pd.read_csv(sfile,usecols=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
blog = pd.read_csv(bfile,index_col=False,usecols=[0,1,2,3,4,5,6,7,8])
clog = pd.read_csv(cfile,index_col=False,usecols=[0,1,2,3,4,5])
# crlog = pd.read_csv(crfile,index_col=False,usecols=[0,1,2,3,4,5])
plog = pd.read_csv(pfile,index_col=False,usecols=[0,1,2,3,4,5])


#Name columns
slog.columns = ['DATE','TIME','OPERATOR_NAME', 'PART_NUMBER','ACCOUNT_NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION','REPRODUCED','SEQUENCE_NUMBER','BUNDLE_NUMBER','CASE_NUMBER','PALLET_NUMBER','CAM1','CAM2','CAM3','CAM4','CAM5','CAM6','BARCODE_GRADE']
blog.columns = ['DATE','TIME','BUNDLE NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION','REPRODUCED','WEIGHING SCALE', 'BARCODE VERIFIER']
clog.columns = ['DATE','TIME','CASE NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION']
# crlog.columns = ['DATE','TIME','CASE NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION']
plog.columns = ['DATE','TIME','PALLET NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION']

#Create workbook to extract data
workbook  = xlsxwriter.Workbook('Log Analysis.xlsx')
sheet = workbook.add_worksheet('Yield')

#Combine Date and Time columns to creat standard date/time format
starttime = datetime.strptime(slog.DATE[0] + ' ' + slog.TIME[0], '%d/%m/%Y %I:%M:%S %p')
sendtime = datetime.strptime(slog.DATE[slog.shape[0] - 1] + ' ' + slog.TIME[slog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
bendtime = datetime.strptime(blog.DATE[blog.shape[0] - 1] + ' ' + blog.TIME[blog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
cendtime = datetime.strptime(clog.DATE[clog.shape[0] - 1] + ' ' + clog.TIME[clog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
# crendtime = datetime.strptime(crlog.DATE[crlog.shape[0] - 1] + ' ' + crlog.TIME[crlog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')

#Pallet log is only file that may be empty
if(plog.shape[0] > 1):
    pendtime = datetime.strptime(plog.DATE[plog.shape[0] - 1] + ' ' + plog.TIME[plog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')

#Extract start and end times for sleeve and bundle logs
times = [starttime,sendtime,bendtime,cendtime]

print('Start Date: '+ str(starttime))

completedate = max(times)
print('Complete Date: ' + str(completedate))

elapsed = max(times)-min(times)
print('Total Production Hrs: ' + str(elapsed))


sleevepn = slog.PART_NUMBER[0]
print('Sleeve Part Number: ' + sleevepn)

runqty = slog.shape[0]
print('Total Run Qty: '+ str(runqty))

#def production_stats(logfile,column,value,condition,message):
#    logfile[logfile.column condition value].shape[0]

produced = slog[slog.TAGGED != 'TAGGED'].shape[0] 
print('Total Produced: ' + str(produced))

tagged = slog[slog.TAGGED == 'TAGGED'].shape[0]
print('Total Tagged: ' + str(tagged))

inserts = slog[slog.TAG_REASON == 'CI2'].shape[0]
print('Defective Inserts: ' + str(inserts))

sleeves = slog[slog.TAG_REASON == 'CI1'].shape[0]
print('Defective Sleeves: ' + str(sleeves))

firstpass = slog[slog.REPRODUCED != 'RE-PRODUCED'] #total run qty in first pass
firstgood = firstpass[firstpass.TAGGED != 'TAGGED'].shape[0]
print('First Pass Yield: ' + str(firstgood))

firstbad = firstpass[firstpass.TAGGED == 'TAGGED'].shape[0]
print('First Pass Defects: ' + str(firstbad))

secondpass = slog[slog.REPRODUCED == 'RE-PRODUCED']
secondgood = secondpass[secondpass.TAGGED != 'TAGGED'].shape[0]
print('Second Pass Yield: ' + str(secondgood))

secondbad = secondpass[secondpass.TAGGED == 'TAGGED'].shape[0]
print('Second Pass Defects: ' + str(secondbad))

sleeverework = slog[slog.REPRODUCED == 'RE-PRODUCED']
sleeverework = sleeverework[sleeverework.TAGGED != 'TAGGED'].shape[0]
print('Sleeve Rework Qty: ' + str(sleeverework))

bundlerework = blog[blog.REPRODUCED == 'RE-PRODUCED']
bundlerework = bundlerework[bundlerework.TAGGED != 'TAGGED'].shape[0]
print('Bundle Rework Qty: ' + str(bundlerework) + '\n')


#--------------DOWNTIME----------------
# downtimesheet= workbook.add_worksheet('Downtime')
# downtimesheet.write(0,0,'From')
# downtimesheet.write(0,1,'To')
# downtimesheet.write(0,2,'Downtime')

# downtimerow = 1
# for i in range(slog.TIME.shape[0]):
#     threshold = datetime.strptime('10:23:00 PM', '%I:%M:%S %p')-datetime.strptime('10:22:30 PM', '%I:%M:%S %p') 
#     if i >0:
#         time1 = datetime.strptime(slog.TIME[i-1], '%I:%M:%S %p')
#         time2 = datetime.strptime(slog.TIME[i], '%I:%M:%S %p')
#         diff = time2-time1

#         if diff > threshold:
#             # print(diff)
#             # print('From '+ str(time1) +' To: ' + str(time2))
#             downtimesheet.write(downtimerow,0,str(time1))
#             downtimesheet.write(downtimerow,1,str(time2))
#             downtimesheet.write(downtimerow,2,str(diff))
#             downtimerow += 1            


#-----------------SLEEVE TAGS------------------------------

#create new df with count of all sleeve tag reasons 
tags = slog.TAG_REASON.value_counts().reset_index(inplace=False)
tags.columns = ['Tag_Reason', 'Count']

#create list of percent tagged (each tag reason / total tagged)
defectpercent = []
for i in tags.Count:
    defectpercent.append((i/tagged)*100)

#add list of defect percentages as a column to df
tags['percent'] = defectpercent 
print(tags)


#Write data to excel workbook
desc = ['Start Date','Complete Date','Total Production Hrs','Sleeve Part Number','Total Run Qty','Total Produced',
        'Total Tagged','Defective Inserts','Defective Sleeves','First Pass Yield','First Pass Defects','Second Pass Yield',
        'Second Pass Defects','Sleeve Rework Qty','Bundle Rework Qty']
data = [starttime,completedate,elapsed,sleevepn,runqty,produced,tagged,inserts,sleeves,firstgood,firstbad,secondgood,secondbad,sleeverework,bundlerework]

for i in range(len(desc)):
    sheet.write(i,0,desc[i])
    sheet.write(i,1,str(data[i]))

sheet.set_column(0,0,18)
sheet.write(len(desc)+2,0,'Tag Reason',workbook.add_format({'bold': True}))
sheet.write(len(desc)+2,1,'Count',workbook.add_format({'bold': True}))
sheet.write(len(desc)+2,2,'%',workbook.add_format({'bold': True}))
    
for i in range(tags.shape[0]):
    sheet.write(len(desc)+3+i,0,tags.Tag_Reason[i])
    sheet.write(len(desc)+3+i,1,tags.Count[i])
    sheet.write(len(desc)+3+i,2,tags.percent[i])


#-----------------------BUNDLE TAGS---------------------

#create new df with count of all bundle tag reasons
btags = blog.TAG_DESCRIPTION.value_counts().reset_index(inplace=False)
btags.columns = ['Tag_Reason', 'Count']
print(btags)

#write to excel workbook
sheet.write(len(desc)+4+tags.shape[0],0,'Bundle Tag Reason',workbook.add_format({'bold': True}))
sheet.write(len(desc)+4+tags.shape[0],1,'Count', workbook.add_format({'bold': True}))

for i in range(btags.shape[0]):
    sheet.write(len(desc)+5+i+tags.shape[0],0,btags.Tag_Reason[i])
    sheet.write(len(desc)+5+i+tags.shape[0],1,btags.Count[i])


#-------------------Process Capabilities----------------------------------
sheet2 = workbook.add_worksheet('CPk')

sheet2.set_column(0,0,19)
sheet2.write(0,1,'CAM3')
sheet2.write(0,2,'CAM4')
sheet2.write(0,3,'CAM5')
sheet2.write(0,4,'CAM6')

def sheet2write(x,y): 
    sheet2.write(x,y,'Max')
    sheet2.write(x+1,y,'Min')
    sheet2.write(x+2,y,'Standard Deviation')
    sheet2.write(x+3,y,'USL')
    sheet2.write(x+4,y,'LSL')
    sheet2.write(x+5,y,'Cp')
    sheet2.write(x+6,y,'Cpk')

sheet2write(1, 0)
sheet2write(9, 0)
sheet2write(17, 0)
sheet2write(25, 0)
sheet2write(33, 0)
sheet2write(41, 0)
sheet2write(49, 0)
sheet2write(57, 0)
sheet2write(65, 0)


def filter_data(column):
    
    #Split string in column by '|' character (creates a list)
    camdf = pd.DataFrame(slog[column].dropna().str.split('|').tolist())
    
    #create empty "filtered dataframe"
    fltrd_camdf = pd.DataFrame()
    
    #iterate through each column of camdf
    for i in range(camdf.shape[1]):
        if('[' in camdf.iloc[0,i]):
            
            #if column contains '[' take values up to ':' and add column to filtered df 
            fltrd_camdf[fltrd_camdf.shape[1]] = camdf[camdf.columns[i]].apply(lambda x: x.split(':')[0]) 
            
    return fltrd_camdf  


#calculate standard deviation of each dataframe column
def std_dev(cam):
    std = []
    for i in range(cam.shape[1]):
        #convert ith column to np array
        column = np.array(cam.iloc[:,i],dtype=float)
        std.append(np.nanstd(column))
    return std

#calculate mean of each dataframe column
def mean(cam):
    mean = []
    for i in range(cam.shape[1]):
        #convert ith column to np array 
        column = np.array(cam.iloc[:,[i]],dtype=float)
        mean.append(np.mean(column))   
    return mean

#find max & min in each dataframe column
def max_min(cam):  
    maxi,mini = [],[]
    for i in range(cam.shape[1]):
        column = np.array(cam.iloc[:,[i]],dtype=float)
        maxi.append(np.amax(column))
        mini.append(np.amin(column))
        # max_min = [np.amin(column,axis=0),np.amax(column,axis=0)]

    return mini,maxi

#calculate cp and cpk of each dataframe column
def Cp(cam,camnumber):   
    stdlist = std_dev(cam)
    mnlist = mean(cam)
    cp,cpl,cpu,cpk,usllist,lsllist=[],[],[],[],[],[]
    
    for i in range(len(stdlist)):
        
        if(camnumber == 3):
            if(i<=5):
                lsl = 100.0
                usl = 5000.0 
            else:
                lsl = 0.0
                usl = 99999.0  
        
        elif(camnumber == 4):
            lsl = 190000
            usl = 262000 
        
        elif(camnumber == 5):
            if(i <= 1):
                lsl = 820.0
                usl = 840.0
            elif(i==2):
                lsl = 10000.0
                usl = 45000.0
            elif(i==3 or i==4):
                lsl = 950
                usl = 1150
            elif(i==5):
                lsl = 1000
                usl = 20000
            elif(i==6 or i==7):
                lsl = 1100
                usl = 1300
            elif(i==8):
                lsl = 500
                usl = 3000
        elif(camnumber ==6):
            if(i <= 1):
                lsl = 890.0
                usl = 930.0
            elif(i==2):
                lsl = 120000.0
                usl = 130000.0
        
        lsllist.append(lsl)
        usllist.append(usl)
        cp.append((usl-lsl)/(6*stdlist[i]))
        cpl.append((mnlist[i]-lsl)/(3*stdlist[i]))
        cpu.append((usl-mnlist[i])/(3*stdlist[i]))  
        cpk.append(min(cpl[i],cpu[i]))
  
    values = {'cp':cp,'cpk':cpk,'std':stdlist,'usl':usllist,'lsl':lsllist,'mean':mnlist,'cpl':cpl,'cpu':cpu}
    return values

def write_cp(cp,max_min,camnumber):
    column_map = {3:1,4:2,5:3,6:4}
    col = column_map[camnumber]
    r=1
    
    for i in range(len(cp['cp'])):
        sheet2.write(r,col,max_min[0][i])
        sheet2.write(1+r,col,max_min[1][i])
        sheet2.write(2+r,col,cp['std'][i])
        sheet2.write(3+r,col,cp['usl'][i])
        sheet2.write(4+r,col,cp['lsl'][i])
        sheet2.write(5+r,col,cp['cp'][i])
        sheet2.write(6+r,col,cp['cpk'][i])      
        r+=8

print('\nAnalyzing CAM3 CPk...')
CAM3df = filter_data('CAM3')
print(max_min(CAM3df))
print(Cp(CAM3df,3))
write_cp(Cp(CAM3df,3),max_min(CAM3df),3)


print('\nAnalyzing CAM4 CPk...')       
CAM4df = filter_data('CAM4')
print(max_min(CAM4df))
print(Cp(CAM4df,4))
write_cp(Cp(CAM4df,4),max_min(CAM4df),4)


print('\nAnalyzing CAM5 CPk...')
CAM5df = filter_data('CAM5')
print(max_min(CAM5df))
print(Cp(CAM5df,5))
write_cp(Cp(CAM5df,5),max_min(CAM5df),5)

    
print('\nAnalyzing CAM6 CPk...')
CAM6df = filter_data('CAM6')
print(max_min(CAM6df))
print(Cp(CAM6df,6))
write_cp(Cp(CAM6df,6),max_min(CAM6df),6)


#---------------------Barcdode Grade----------------------------------------------------
print('\nAnalyzing Barcode Grade...')

sheet3 = workbook.add_worksheet('Barcode Grade')

BARCODE_GRADE = slog.BARCODE_GRADE.dropna()
BARCODE_GRADE = BARCODE_GRADE.reset_index(drop=True)
for i in range(BARCODE_GRADE.shape[0]):
    if(BARCODE_GRADE[i] != '0'):
        BARCODE_GRADE[i] = float(BARCODE_GRADE[i][1:])
    else:
        BARCODE_GRADE[i] = float(BARCODE_GRADE[i])

A,B,C,D,F = 0,0,0,0,0
for i in BARCODE_GRADE: 
    if(i<0.5):
        F += 1
    elif(0.5 <= i <= 1.5):
        D += 1
    elif(1.5 < i <= 2.5):
        C += 1
    elif(2.5 < i <= 3.5):
        B += 1
    elif(3.5 < i <= 4):
        A += 1
        
maxgrade = max(BARCODE_GRADE)
mingrade = min(BARCODE_GRADE)
avggrade = np.mean(BARCODE_GRADE)

print('\nMax: ' + str(maxgrade))
print('Min: ' + str(mingrade))
print('Average Grade: ' + str(avggrade))

sheet3.set_column(0,0,15)
sheet3.write(0,0, 'Max Grade')
sheet3.write(1,0, 'Min Grade')
sheet3.write(2,0, 'Average Grade')
sheet3.write(0,1, maxgrade)
sheet3.write(1,1, mingrade)
sheet3.write(2,1, avggrade)
sheet3.write(4,0, 'Grade')
sheet3.write(4,1, 'Count')
sheet3.write(5,0, 'A')
sheet3.write(6,0, 'B')
sheet3.write(7,0, 'C')
sheet3.write(8,0, 'D')
sheet3.write(9,0, 'F')
sheet3.write(5,1, A)
sheet3.write(6,1, B)
sheet3.write(7,1, C)
sheet3.write(8,1, D)
sheet3.write(9,1, F)
  
workbook.close()

k=input('\nPress Enter to Exit')
