import os
import pandas as pd
#not used
#cwd = os.getcwd()

#os.chdir('')
#--------

#to do: make this a dynamic file read in
file = 'fundlll.xlsx'
spread_sheet = pd.ExcelFile(file) 

#reads in the excel file as an DataFrame object. Sheet_names is a list of all sheet names within the Excel file.
sheet1 = spread_sheet.parse('Comp Fund Stats')
sheet2 = spread_sheet.parse('Fund I Overview')
sheet3 = spread_sheet.parse('Fund II Overview')

#to get data items from the sheet.
#By updating this list, you can pull more data items
s1_variablelist = ['# Investments', 'Avg Initial Check Size (led deals)', 'Avg Round Size (lead)']


#get all of the index values correpsonding to the s1_variablelist. 
our_list=[]
sheet_data = pd.DataFrame()
for x in sheet1.loc[:,'Comparative Fund Statistics']:
    if x in s1_variablelist:
        labels = sheet1[sheet1['Comparative Fund Statistics'] == x].index.values.tolist()
        our_list.append(labels)
#-----------------------------

our_list = [our_list[x][0] for x in range(len(our_list))]
sheet_data = pd.DataFrame()
x = 0
for x in range(sheet1.shape[0]):
    if x in our_list:
        pass
    else:
        sheet1 = sheet1.drop(x)
#------------------------------------------------------------

print(sheet1)



















