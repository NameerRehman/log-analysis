# -*- coding: utf-8 -*-
import pandas as pd
import xlsxwriter
from datetime import datetime
import numpy as np

# slogin = input("Enter directory of Sleeve log: ")

#/Users/namee/Desktop/FAT Reports/FAT_C1_40(25_3) + HEADER/PS FAT 1_C1_40_39600_S_LOG.csv')
# /Users/namee/Desktop/FAT Reports/FAT_C1_40(25_3) + HEADER/PS FAT 1_C1_40_39600_B_LOG.csv',index_col=False)
# /Users/namee/Desktop/FAT Reports/FAT_C1_40(25_3) + HEADER/PS FAT 1_C1_40_39600_C_LOG.csv',index_col=False)
# crlog = pd.read_csv('/Users/namee/Desktop/FAT Reports/FAT_C1_40(25_3) + HEADER/PS FAT 1_C1_40_39600_CR_LOG.csv',index_col=False)
# plog = pd.read_csv('/Users/namee/Desktop/FAT Reports/FAT_C1_40(25_3) + HEADER/PS FAT 1_C1_40_39600_P_LOG.csv',index_col=False)


sfile = input('Enter path of S log: ')
bfile = input('Enter path of B log: ')
cfile = input('Enter path of C log: ')
crfile = input('Enter path of CR log: ')
pfile = input('Enter path of P log: ')

print('\nAnalyzing log files, please wait...')

slog = pd.read_csv(sfile,usecols=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19])
blog = pd.read_csv(bfile,index_col=False,usecols=[0,1,2,3,4,5,6,7,8])
clog = pd.read_csv(cfile,index_col=False,usecols=[0,1,2,3,4,5])
crlog = pd.read_csv(crfile,index_col=False,usecols=[0,1,2,3,4,5])
plog = pd.read_csv(pfile,index_col=False,usecols=[0,1,2,3,4,5])

#slog = slog.dropna(axis=1,how='all')
# blog = blog.dropna(axis=1,how='all')
# clog = blog.dropna(axis=1,how='all')
# crlog = blog.dropna(axis=1,how='all')

workbook  = xlsxwriter.Workbook('/Users/namee/desktop/filename1.xlsx')
sheet = workbook.add_worksheet('Data')

slog.columns = ['DATE','TIME','OPERATOR_NAME', 'PART_NUMBER','ACCOUNT_NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION','REPRODUCED','SEQUENCE_NUMBER','BUNDLE_NUMBER','CASE_NUMBER','PALLET_NUMBER','CAM1','CAM2','CAM3','CAM4','CAM5','CAM6','BARCODE_GRADE']
blog.columns = ['DATE','TIME','BUNDLE NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION','REPRODUCED','WEIGHING SCALE', 'BARCODE VERIFIER']
clog.columns = ['DATE','TIME','CASE NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION']
crlog.columns = ['DATE','TIME','CASE NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION']
plog.columns = ['DATE','TIME','PALLET NUMBER','TAGGED','TAG_REASON','TAG_DESCRIPTION']


slog.to_excel('slog2.xlsx')

startdate = slog.DATE[0]
startyear = startdate[6:9]
startmonth = startdate[3:4]

starttime = datetime.strptime(slog.DATE[0] + ' ' + slog.TIME[0], '%d/%m/%Y %I:%M:%S %p')
sendtime = datetime.strptime(slog.DATE[slog.shape[0] - 1] + ' ' + slog.TIME[slog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
bendtime = datetime.strptime(blog.DATE[blog.shape[0] - 1] + ' ' + blog.TIME[blog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
cendtime = datetime.strptime(clog.DATE[clog.shape[0] - 1] + ' ' + clog.TIME[clog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
crendtime = datetime.strptime(crlog.DATE[crlog.shape[0] - 1] + ' ' + crlog.TIME[crlog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')
if(plog.shape[0] > 1):
    pendtime = datetime.strptime(plog.DATE[plog.shape[0] - 1] + ' ' + plog.TIME[plog.shape[0] - 1], '%d/%m/%Y %I:%M:%S %p')

times = [starttime,sendtime,bendtime,cendtime,crendtime]

print('Start Date: '+ str(starttime))

completedate = max(times)
print('Complete Date: ' + str(completedate))

elapsed = max(times)-min(times)
print('Total Production Hrs: ' + str(elapsed))

sleevepn = slog.PART_NUMBER[0]
print('Sleeve Part Number: ' + sleevepn)

runqty = slog.shape[0]
print('Total Run Qty: '+ str(runqty))

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
secondgood = secondpass[secondpass.TAGGED != 'TAGGED']
print('Second Pass Yield: ' + str(secondgood))

secondbad = secondpass[secondpass.TAGGED == 'TAGGED'].shape[0]
print('Second Pass Defects: ' + str(secondbad))

sleeverework = slog[slog.REPRODUCED == 'RE-PRODUCED']
sleeverework = sleeverework[sleeverework.TAGGED != 'TAGGED'].shape[0]
print('Sleeve Rework Qty: ' + str(sleeverework))

bundlerework = blog[blog.REPRODUCED == 'RE-PRODUCED']
bundlerework = bundlerework[bundlerework.TAGGED != 'TAGGED'].shape[0]
print('Bundle Rework Qty: ' + str(bundlerework) + '\n')

tags = slog.TAG_REASON.value_counts().reset_index(inplace=False)
tags.columns = ['Tag_Reason', 'Count']

defpercent = []
for i in tags.Count:
    defpercent.append((i/tagged)*100)

tags['percent'] = defpercent #add list of defect percentages as a column to df
print(tags)

desc = ['Start Date','Complete Date','Total Production Hrs','Sleeve Part Number','Total Run Qty','Total Produced',
        'Total Tagged','Defective Inserts','Defective Sleeves','First Pass Yield','First Pass Defects','Second Pass Yield'
        'Second Pass Defects','Sleeve Rework Qty','Bundle Rework Qty']
analysis = [starttime,completedate,elapsed,sleevepn,runqty,produced,tagged,inserts,sleeves,firstgood,firstbad,secondgood,secondbad,sleeverework,bundlerework]

for i in range(len(desc)):
    sheet.write(i,0,desc[i])
    sheet.write(i,1,str(analysis[i]))

sheet.write(len(desc)+2,0,'Tag Reason')
sheet.write(len(desc)+2,1,'Count')
sheet.write(len(desc)+2,2,'%')
    
for i in range(tags.shape[0]):
    sheet.write(len(desc)+3+i,0,tags.Tag_Reason[i])
    sheet.write(len(desc)+3+i,1,tags.Count[i])
    sheet.write(len(desc)+3+i,2,tags.percent[i])

workbook.close()
    
#---------------------CAM3 Glue----------------------------------------------------
print('\nAnalyzing CAM3 CPk...')

CAM3 = slog.CAM3.str.split('|').dropna()
CAM3 = CAM3.reset_index(drop=True)

for i in range(CAM3.shape[0]):
    CAM3[i] = [j for j in CAM3[i] if '[' in j] #Change row to new list without unwanted entries
    for j in range(len(CAM3[0])):
        CAM3[i][j] = CAM3[i][j][: CAM3[i][j].find(':')] #only take text up to the : character

list = []
for h in range(len(CAM3[0])):
    for i in range(CAM3.shape[0]): 
        list.append(CAM3[i][h]) 

    list= np.array(list,dtype=float)
    print(np.std(list,axis=0,ddof=0))
    
    list=[]
    
#---------------------CAM4 Insert----------------------------------------------------
print('\nAnalyzing CAM4 CPk...')

CAM4 = slog.CAM4.str.split('|').dropna()
CAM4 = CAM4.reset_index(drop=True)
for i in range(CAM4.shape[0]):
    CAM4[i] = [j for j in CAM4[i] if '[' in j] #Change row to new list without unwanted entries
    for j in range(len(CAM4[0])):
        CAM4[i][j] = CAM4[i][j][: CAM4[i][j].find(':')] #only take text up to the : character

list1 = []
for h in range(len(CAM4[0])):
    for i in range(CAM4.shape[0]): 
        list1.append(CAM4[i][h]) 

    list1= np.array(list1,dtype=float)
    
    lsl4 = 190000
    usl4 = 262000 
        
    std4=np.std(list1)
    mn4 = np.mean(list1,axis=0)
    Cp4 = (usl4-lsl4)/(6*std4)
    Cpl4 = (mn4 - lsl4)/(3*std4)
    Cpu4 = (usl4 - mn4)/(3*std4)
    Cpk4 = min(Cpl4,Cpu4)

    print('\nMax: ' + str(max(list1)))
    print('Min: ' + str(min(list1)))
    print('Standard Deviation: ' + str(std4))
    print('USL: ' + str(usl4))
    print('LSL: ' + str(lsl4))
    print('Cp: ' + str(Cp4))
    print('Cpk: ' + str(Cpk4))

    list1=[]

#---------------------CAM5 Rip Cord/Sleeve----------------------------------------------------
print('\nAnalyzing CAM5 CPk...')

CAM5 = slog.CAM5.str.split('|').dropna()
CAM5 = CAM5.reset_index(drop=True)
for i in range(CAM5.shape[0]):
    CAM5[i] = [j for j in CAM5[i] if '[' in j] #Change row to new list without unwanted entries
    for j in range(len(CAM5[0])):
        CAM5[i][j] = CAM5[i][j][: CAM5[i][j].find(':')] #only take text up to the : character

list1 = []
for h in range(len(CAM5[0])):
    for i in range(CAM5.shape[0]): 
        num = float(CAM5[i][h])
        if(num > 0 and num != 99999.0 and num != 999000.0):
            list1.append(num) 

    list1= np.array(list1)
    
    if(h==0):
        lsl5 = 820.0
        usl5 = 840.0
        window = '1'
    elif(h==1):
        lsl5 = 820.0
        usl5 = 840.0 
        window = '2'
    elif(h==2):
        lsl5 = 10000.0
        usl5 = 45000.0
        window = '3'
    elif(h==3):
        lsl5 = 950
        usl5 = 1150
        window = '4'
    elif(h==4):
        lsl5 = 950
        usl5 = 1150
        window = '5'
    elif(h==5):
        lsl5 = 1000
        usl5 = 20000
        window = '6'
    elif(h==6):
        lsl5 = 1100
        usl5 = 1300
        window = '7'
    elif(h==7):
        lsl5 = 1100
        usl5 = 1300
        window = '8'
    elif(h==8):
        lsl5 = 500
        usl5 = 3000
        window = '9'
        
    std5=np.std(list1)
    mn5 = np.mean(list1,axis=0)
    Cp5 = (usl5-lsl5)/(6*std5)
    Cpl5 = (mn5 - lsl5)/(3*std5)
    Cpu5 = (usl5 - mn5)/(3*std5)
    Cpk5 = min(Cpl5,Cpu5)

    print('\nW' + window + ' Max: ' + str(max(list1)))
    print('W' + window + ' Min: ' + str(min(list1)))
    print('W' + window + ' Standard Deviation: ' + str(std5))
    print('W' + window + ' USL: ' + str(usl5))
    print('W' + window + ' LSL: ' + str(lsl5))
    print('W' + window + ' Cp: ' + str(Cp5))
    print('W' + window + ' Cpk: ' + str(Cpk5))
    
    list1=[]

#---------------------CAM6 Barcode----------------------------------------------------
print('\nAnalyzing CAM6 CPk...')

CAM6 = slog.CAM6.str.split('|').dropna()
CAM6 = CAM6.reset_index(drop=True)
for i in range(CAM6.shape[0]):
    CAM6[i] = [j for j in CAM6[i] if '[' in j] #Change row to new list without unwanted entries
    for j in range(len(CAM6[0])):
        CAM6[i][j] = CAM6[i][j][: CAM6[i][j].find(':')] #only take text up to the : character

list1 = []
for h in range(len(CAM6[0])):
    for i in range(CAM6.shape[0]):
        num = float(CAM6[i][h])
        list1.append(num) 

    list1= np.array(list1)

    if(h==0):
        lsl6 = 890.0
        usl6 = 930.0
        window = '1'
    elif(h==1):
        lsl6 = 890.0
        usl6 = 930.0 
        window = '2'
    elif(h==2):
        lsl6 = 120000.0
        usl6 = 130000.0
        window = '3'
        
    std6=np.std(list1)
    mn6 = np.mean(list1,axis=0)
    Cp6 = (usl6-lsl5)/(6*std6)
    Cpl6 = (mn6 - lsl5)/(3*std6)
    Cpu6 = (usl6 - mn6)/(3*std6)
    Cpk6 = min(Cpl6,Cpu6)

    print('\nW' + window + ' Max: ' + str(max(list1)))
    print('W' + window + ' Min: ' + str(min(list1)))
    print('W' + window + ' Standard Deviation: ' + str(std6))
    print('W' + window + ' USL: ' + str(usl6))
    print('W' + window + ' LSL: ' + str(lsl6))
    print('W' + window + ' Cp: ' + str(Cp6))
    print('W' + window + ' Cpk: ' + str(Cpk6))
     
    list1=[]
    
#---------------------Barcdode Grade----------------------------------------------------
print('\nAnalyzing Barcode Grade...')

BARCODE_GRADE = slog.BARCODE_GRADE.dropna()
BARCODE_GRADE = BARCODE_GRADE.reset_index(drop=True)
for i in range(BARCODE_GRADE.shape[0]):
    if(BARCODE_GRADE[i] != '0'):
        BARCODE_GRADE[i] = float(BARCODE_GRADE[i][1:])
    else:
        BARCODE_GRADE[i] = float(BARCODE_GRADE[i])
        
maxgrade = max(BARCODE_GRADE)
mingrade = min(BARCODE_GRADE)
avggrade = np.mean(BARCODE_GRADE)

print('\nMax: ' + str(maxgrade))
print('Min: ' + str(mingrade))
print('Average Grade: ' + str(avggrade))

k=input('\nPress Enter to Exit')