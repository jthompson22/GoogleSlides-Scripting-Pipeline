import os
import pandas as pd
import numpy 



# This script is used to read in the necessary data from an 'xlsx' file, and pass it to Main 
# where it will be uploaded to Google Drive

def Get_Data():
    #to do: make this a dynamic file read in
    file = 'jt.xlsx'
    spread_sheet = pd.ExcelFile(file) 
    

    #reads in the excel file as an DataFrame object. Sheet_names is a list of all sheet names within the Excel file.
    #spread_sheet.columns = spread_sheet.columns.str.replace(' ', '_')
    sheet1 = spread_sheet.parse('transferData')
    #sheet1.columns = ['tagName', 'value']
    
    #sheet1.columns = ['Comparative Fund Statistics', 'Fund 1', 'Fund 2', 'Fund 3', 'misc4', 'misc5', 'misc6', \
                   # 'misc7', 'misc8', 'misc9', 'misc10']
    #sheet2 = spread_sheet.parse('Fund I Overview')
    #sheet2.columns = ['A', 'B']
    #sheet3 = spread_sheet.parse('Fund II Overview')
    #sheet3.columns = ['A', 'B']

    # Each list here contains the the item labels we want from the excel. By updating this list, you update what is being pulled from the excel doc.
    # To add a data lablel:
        #
    #s1_variablelist = ['# Investments', 'Avg Initial Check Size (led deals)', 'Avg Round Size (le]
    #s2_variablelist = ['Gross ROIC', 'Net TVPI', 'DPI (+ PheonixDist)', 'Seed to Series A Grad Rate', 'Net IRR']

    #-----------------
    #return all variables we need from Comparative Fund Statistics.
    # 
    # 

    #sheet1_var_list=[]
    #for x in sheet1['Comparative Fund Statistics']:
    #    if x in s1_variablelist:
    #        labels = sheet1[sheet1['Comparative Fund Statistics'] == x].index.values.tolist()
    #        sheet1_var_list.append(labels)
        #update Sheet 1 dataDataframe
    #sheet1_var_list = [sheet1_var_list[x][0] for x in range(len(sheet1_var_list))]
    #for x in range(sheet1.shape[0]):
    #   if x in sheet1_var_list:
    #        pass
    #    else:
    #        sheet1 = sheet1.drop(x)
    
    #print(sheet1)
    #return all variables we need from Fund1 1 Overview, or sheet2 

    #sheet2_var_list=[]
    #for x in sheet2['B']:
    #   if x in s2_variablelist:
    #        labels = sheet2[sheet2['B'] == x].index.values.tolist()
    #        sheet2_var_list.append(labels)
    
    #sheet2_var_list = [sheet2_var_list[x][0] for x in range(len(sheet2_var_list))]
    #sheet2 var list will now contain the index of every point we need
    #for x in range(sheet2.shape[0]):
    #    if x in sheet2_var_list:
    #        pass
    #    else:
    #        sheet2 = sheet2.drop(x)
    
    #-----------------------------ls


   
    #------------------------------------------------------------

    #value1, value2, value3 = [x for x in sheet1['Fund 1']]
    #print(value1, value2, value3)
    # print(Fund1_values)
    #print('-------')
    #print('Sheet 2 %s' % sheet2)'''
    #print(sheet1)
    #our_list = [x for x in sheet1['1']]
    #print(our_list)
    ser = pd.Series(sheet1['tagName']) 
    values = pd.Series(sheet1['value'])
    our_type = pd.Series(sheet1['type'])
    print(our_type)
    print(ser.shape[0])
    print(ser[0])
    print(values)
    print(ser.iloc[0])
    items = {}
    for x in range(ser.shape[0]):
        if pd.isnull(ser.iloc[x]): 
            break
        else:
            if our_type[x] == 'none':
                if '.0' in str(values[x]):
                    y = str("{:,.0f}".format(values[x]))
                    items[ser[x]] = y
                else:
                    items[ser[x]] = str(values[x])
            elif our_type[x] == 'percent':
                if '.' in str(values[x]):
                    y = str("{:.2f}%".format(values[x]))
                    print(y)
                    items[ser[x]] = y
                else:
                    items[ser[x]] = str(values[x]) + "%"
            elif our_type[x] == 'money':
                if '.0' in str(values[x]):
                    y = str("${:,.0f}".format(values[x]))
                    items[ser[x]] = y
                else:
                    items[ser[x]] = str("${:,}".format(values[x]))
                
            #print(ser[x])
            #items[ser[x]] = values[x]
            

    print(items)

    #print(len(items))
    #print('hello')
    
    #print(sheet1.loc[1]['tagName'])
    
    #print(sheet1['tagName'])
    #print(sheet2)

   # return sheet1, sheet2


if __name__ == '__main__':
    Get_Data()
















