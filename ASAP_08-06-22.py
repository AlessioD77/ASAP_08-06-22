# -*- coding: utf-8 -*-
"""
Created on Tue Mar 22 15:10:05 2022

Automated script to automatically calculate site-specific N-glycan composition

Takes 3 inputs:
    raw data CSV
    PNGS csv
    Glycan_S file (stored in /ASAP/)

@author: alekoj
"""

import pandas as pd
import numpy as np

raw_df_filename = input('Enter the raw data filename: ') + '.csv'
final_filename = raw_df_filename[:-4] + '.xlsx'
pngs_filename = input('Enter the PNGS data filename: ') + '.csv'
try:
    raw_df = pd.read_csv(raw_df_filename)
    pngs = pd.read_csv(pngs_filename)
except:
    print('File cannot be opened: ', raw_df_filename)    
    exit()

classification_df=pd.read_csv("ASAP/Glycan_s.csv")


filter_df = raw_df[['Start\r\nAA', 'End\r\nAA', 'Var. Pos.\r\nProtein', 'Sequence', 'Glycans', 'Validate', 'Score', 'XIC area\r\nsummed']]

#make all sequence names upper case
filter_df['Sequence'] = filter_df['Sequence'].str.upper()

for i,rows in filter_df.iterrows():
    #print(rows['Glycans'])
#here is where the main work starts to find the right classfication
    class_name=classification_df[classification_df['Glycans'] == rows['Glycans']]
    #print(class_name['Species'])
#then its just a matter of assigning the values
    try:
        filter_df.at[i,'Classification']=class_name.iloc[0]['Species']
    except:
        filter_df.at[i,'Classification']='Unoccupied'
 
        #Code for assigning PNGS, remove below up to line 83 to not do PNGSs.
png_dict = dict(pngs[['Pos', 'Hxb2 Env coords']].values)

values = []
for i, row in filter_df.iterrows():
    print(row['Var. Pos.\r\nProtein'])
    sp = str(row['Var. Pos.\r\nProtein']).split(',')
    print(sp)
    if sp[0] == '' or sp[0] == 'nan':
        val = np.NaN
    elif len(sp)>1:
        num1, num2 = sp
        print(num1, num2)
        try:
            val = png_dict[int(num1)]
        except KeyError:
            try:
                val = png_dict[int(num2)]
            except KeyError:
                val = np.NaN

    else:
       
        num = int(float(sp[0])) #this seems to crash the code with some CSVs for some reason. 03-22-22(ValueError: invalid literal for int() with base 10: '122.0')
        #If the above line gives a ValueError then change to any integer, otherwise num = int(sp[0])
        #UPDATE 03-23-22: num = int(float(sp[0])) seems to work.
        
        print(num)
        try:    
            val = png_dict[int(num)]
        except KeyError or ValueError:
            val = np.NaN
    values.append(val)

filter_df['Hxb2 Env coords'] = values

filter_df = filter_df[['Start\r\nAA', 'End\r\nAA', 'Var. Pos.\r\nProtein', 'Hxb2 Env coords', 'Sequence', 'Glycans', 'Validate', 'Score', 'XIC area\r\nsummed']]

unique_seq = filter_df['Sequence'].unique()

columns = filter_df.columns

writer = pd.ExcelWriter("ASAP/part1.xlsx", engine = 'xlsxwriter')

for col in unique_seq:
    found_df = filter_df[filter_df['Sequence'] == col] #Matching based on column name
    found_df[columns].to_excel(writer, sheet_name=col[:30], index=False)
writer.save()
writer.close()
print('Completed part 1')

part1 = pd.ExcelFile('ASAP/part1.xlsx')
sheet_names = part1.sheet_names
writer = pd.ExcelWriter("ASAP/part2.xlsx", engine = 'xlsxwriter')

for sheet_name in sheet_names:
    print(sheet_name)
    df = pd.read_excel("ASAP/part1.xlsx", sheet_name = sheet_name)
    print(df)
    total = df['XIC area_x000D_\nsummed'].sum() #Getting total XIC area summed for each site
    print(total)

    for i, rows in df.iterrows():
        df.at[i,'Percentage'] = ((rows['XIC area_x000D_\nsummed'] / total )*100)
    df.to_excel(writer, sheet_name = sheet_name, index = False)

writer.save()
writer.close()
print('Completed part 2')

part2 = pd.ExcelFile("ASAP/part2.xlsx")

sheet_names = part2.sheet_names
writer = pd.ExcelWriter("ASAP/part3.xlsx", engine = 'xlsxwriter')


for sheet_name in sheet_names:
    print(sheet_name)
    df = pd.read_excel("ASAP/part2.xlsx", sheet_name = sheet_name)

    for i, rows in df.iterrows():
        print(rows['Glycans'])
        #here is where the main work starts to find the right classfication
        class_name=classification_df[classification_df['Glycans'] == rows['Glycans']]
        print(class_name['Species'])
        #then its just a matter of assigning the values
        try:
          df.at[i,'Classification']=class_name.iloc[0]['Species']
        except:
          df.at[i,'Classification']='Unoccupied'
          # continue
    df.to_excel(writer,sheet_name=sheet_name,index=False)
writer.save()
writer.close()
print("Completed part 3")

raw_df = pd.ExcelFile("ASAP/part3.xlsx")
to_calc = ['M9GLC','M9','M8','M7','M6','M5','M4','M3','FM','HYBRID','FHYBRID','HexNAc(3)(x)',
'HexNAc(3)(F)(x)','HexNAc(4)(x)','HexNAc(4)(F)(x)','HexNAc(5)(x)','HexNAc(5)(F)(x)',
'HexNAc(6+)(x)','HexNAc(6+)(F)(x)','Unoccupied','Core']
sheet_names=raw_df.sheet_names
writer = pd.ExcelWriter("ASAP/part4.xlsx", engine = 'xlsxwriter')

for sheet_name in sheet_names:
    print(sheet_name)
    df=pd.read_excel("ASAP/part3.xlsx",sheet_name=sheet_name)

    for index,calc in enumerate(to_calc):
        #i try to find if those names above(like M8) appear in the table and in which rows
        matched_records=df[df['Classification'] == calc]
        #if i do find 1 or more rows with it then i add them and assign it else its 0
        if(not matched_records.empty):
          value=matched_records['Percentage'].sum()
          df.at[index,'Glycan class']=calc
          df.at[index,'Class percentage']=value
        else:
          df.at[index,'Glycan class']=calc
          df.at[index,'Class percentage']=0.00
    df.to_excel(writer,sheet_name=sheet_name,index=False)

writer.save()
writer.close()
print("Completed part 4")

raw_df=pd.ExcelFile("ASAP/part4.xlsx")
#name of the things to calculate
to_calc=['Oligomannose','Hybrid','Complex','Unoccupied','Core','Fucose','NeuAc/NeuGc']
sheet_names=raw_df.sheet_names
writer = pd.ExcelWriter(final_filename, engine = 'xlsxwriter')

for sheet_name in sheet_names:
    print(sheet_name)
    df=pd.read_excel("ASAP/part4.xlsx",sheet_name=sheet_name)
    #from oligomannose to core the calculation is done based on the row index
    #as i know from which row i need to get the values
    oli=df.iloc[0:9]['Class percentage'].sum()
    df.at[0,'Oligomannose']=oli
    hybrid=df.iloc[9:11]['Class percentage'].sum()
    #print(hybrid)
    df.at[0,'Hybrid']=hybrid
    comple=df.iloc[11:19]['Class percentage'].sum()
    df.at[0,'Complex']=comple
    unocc=df.iloc[19]['Class percentage']
    df.at[0,'Unoccupied']=unocc
    core=df.iloc[20]['Class percentage']
    df.at[0,'Core']=core

    #for Fucose i just try to find "F" in the string and if it exist i add it to total
    total=0
    for i,rows in df.iterrows():
        if(str(rows['Glycan class']).find("F") != -1):
          total+=rows['Class percentage']
    df.at[0,'Fucose']=total

    #for Neu i do the same but with those keywords and an OR operator
    total=0
    for i,rows in df.iterrows():
        if((str(rows['Glycans']).find("NeuAc") != -1) or (str(rows['Glycans']).find("NeuGc") != -1)):
          print(i)
          total+=rows['Percentage']
    df.at[0,'NeuAc/NeuGc']=total
    df.to_excel(writer,sheet_name=sheet_name,index=False)
    
writer.save()
writer.close()
print("Completed Final report: ", final_filename)
