# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

#%%
import pandas as pd
import langid
from openpyxl import load_workbook


def space_counter(name):
    count = 0
    for i in name:
        if (i.isspace()) == True:
            count += 1
    return count 
            




# Open excel sheet 'path'!!!
file = r'C:\Users\work\Desktop\test.xlsx'
new = r'C:\Users\work\Desktop\test1.xlsx'

# Sheet format, 0 = headers on, None = No header 
header_ = 0                     

# Read the sheet

df0 = pd.read_excel(file, header=header_)
col_name = df0.columns[0]


a=[]

for i in range(len(df0[col_name])):
    a.append(df0[col_name][i])
'''
#%%

en=[]
zh=[]
re=[]
for i in a :
    print(langid.classify(i)[0])
    if langid.classify(i)[0] == 'zh':
        zh.append(i)
       # print('#zh =',len(zh))
    else :
        print(i,space_counter(i))
        if space_counter(i) < 2:
            fullname = i.split(' ')
            if len(fullname[0]) < 10 :
                en.append(i)
                print(fullname)
            else:
                re.append(i)
        else:
            re.append(i)
       # print('#en =',len(en))
    
'''



en=[]
zh=[]
re=[]

workbook = load_workbook(filename=file)
#open workbook
sheet = workbook.active


#modify the desired cell
sheet["C1"] = "中文"
sheet["D1"] = "English"
sheet["E1"] = "Revisit"

workbook.save(new)

df = pd.read_excel(new, header=header_)


num_zh = 0
num_en = 0
num_re = 0

for i in a :
    #print(langid.classify(i)[0])
    if langid.classify(i)[0] == 'zh':
        zh.append(i)
        df[df.columns[2]][num_zh] = i
        num_zh += 1
        #print(i,num)
    else :
        if space_counter(i) < 2 :
            fullname = i.split(' ')
            if len(fullname[0]) < 10 :
                en.append(fullname[0])
                df[df.columns[3]][num_en] = fullname[0]
                num_en += 1
                #print(num_en,fullname)
            else:
                re.append(i)
                df[df.columns[4]][num_re] = i
                num_re+=1
                #print('long name') 
        else:
           re.append(i)
           df[df.columns[4]][num_re] = i 
           num_re+=1

df.to_excel(new)