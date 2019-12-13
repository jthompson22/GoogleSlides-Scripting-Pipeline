import os
import pandas as pd
import numpy 
import re
from os import listdir

# This script is used to read in the necessary data from an 'xlsx' file, and pass it to Main 
# where it will be uploaded to Google Drive


#def find_csv_filenames( path_to_dir, suffix=".csv" ):
    #filenames = listdir(path_to_dir)
    #return [ filename for filename in filenames if filename.endswith( suffix ) ]


def Get_Data():
    #to do: make this a dynamic file read in
    dirpath = os.getcwd()

    
    
    file = 'data.xlsx'
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

    '''our_list = ['number-1', 'money-2', 'text-1', 'percent']
    money = 'money-2'
    text = 'text-1'
    percent = 'percent'
    test = 1000000.234
    number = re.search('[0-9]', our_list[1]).group(0)
    test = f'${test:,.{number}f}'
    print(number)
    print(test)'''
    
    ser = pd.Series(sheet1['tagName']) 
    values = pd.Series(sheet1['value'])
    our_type = pd.Series(sheet1['format'])
    items = {}
    for x in range(ser.shape[0]):
        if pd.isnull(ser.iloc[x]): 
            pass
        else:
            if 'number' in our_type[x]:
                format_variable = re.search('[0-9]', our_type[x]).group(0)
                val = f'{values[x]:,.{format_variable}f}'
                val = str(val)
                items[ser[x]] = val
            if 'text' in our_type[x]:
                items[ser[x]] = values[x]
            if 'percent' in our_type[x]:
                format_variable = re.search('[0-9]', our_type[x]).group(0)
                new_value = values[x] * 100
                val = f'{new_value:,.{format_variable}f}%'
                val = str(val)
                items[ser[x]] = val
            if 'money' in our_type[x]:
                format_variable = re.search('[0-9]', our_type[x]).group(0)
                val = f'${values[x]:,.{format_variable}f}'
                val = str(val)
                items[ser[x]] = val
                
    

    '''ser = pd.Series(sheet1['tagName']) 
    values = pd.Series(sheet1['value'])
    our_type = pd.Series(sheet1['format'])
    items = {}
    for x in range(ser.shape[0]):
        if pd.isnull(ser.iloc[x]): 
            pass
        else:
            if our_type[x] == 'none':
                y = str("{:,.0f}".format(values[x]))
                items[ser[x]] = y
            elif our_type[x] == 'none-one':
                y = str("{:,.1f}".format(values[x]))
                items[ser[x]] = y
            elif our_type[x] == 'none-two':
                y = str("{:,.2f}".format(values[x]))
                items[ser[x]] = y
            elif our_type[x] == 'percent':
                #check if percent is in the form .xx (ex, 0.37)
                if '0.' in str(values[x]):
                    y = str("{:.0f}%".format(values[x] * 100))
                    items[ser[x]] = y
                else:
                    y = str("{:.0f}%".format(values[x]))
                    items[ser[x]] = y
            elif our_type[x] == 'percent-one':
                if '0.' in str(values[x]):
                    y = str("{:.1f}%".format(values[x] * 100))
                    items[ser[x]] = y
                else:
                    y = str("{:.1f}%".format(values[x]))
                    items[ser[x]] = y
            elif our_type[x] == 'percent-two':
                if '0.' in str(values[x]):
                    y = str("{:.2f}%".format(values[x] * 100))
                    items[ser[x]] = y
                else:
                    y = str("{:.2f}%".format(values[x]))
                    items[ser[x]] = y
            elif our_type[x] == 'money':
                y = str("${:,.0f}".format(values[x]))
                items[ser[x]] = y
            elif our_type[x] == 'money-two':
                y = str("${:,.2f}".format(values[x]))
                items[ser[x]] = y


    return items'''
    return items
    #for x in items.items():
        #print(x)



if __name__ == '__main__':
    Get_Data()
















