from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload
import datetime
from datetime import date
import os
import sys
import pandas as pd
import numpy 
import re
import json
import time
from urllib.error import HTTPError
import json

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

# The ID of a sample presentation, going to have to read these in dynamically from Pickle object.
master_folder_ID = '1zMrdfLBZdtuGpEq1oE_hGyRkwU9BXARw'
template_presentation_id = '16nmnFPBu4cN8EzRBleeT7oNRbijAhH4cD0NAt2b3dFU'
#archive_folder_ID = drive_id[]

#__________________GET Data Method, gets data from excel sheet named "data" and formats it appropriately to be sent up.

def get_data():
    
    #------read in JSON variables. 

    #excel_document = None
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable) + "/"
        directory_list = os.scandir(application_path)
        for x in directory_list:
            if '.xlsx' in x.name:
                excel_document = application_path + x.name 
        print(x.name)
        print("Executable Running")
    elif __file__:
        print("Script Running")
        application_path = os.path.dirname(__file__) 
        directory_list = os.scandir()
        for x in directory_list:
            if '.xlsx' in x.name:
                excel_document = x.name
        print(x.name)
        print(application_path)

    #make dictionary out of Excel files
    print(excel_document)
    spread_sheet = pd.ExcelFile(excel_document) 
    sheet1 = spread_sheet.parse('transferData')
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
                items["{{" + ser[x] + "}}"] = val
            if 'text' in our_type[x]:
                if (type(values[x]) is datetime.datetime):
                    date = values[x].strftime("%B %d, %Y")
                    items["{{" + ser[x] + "}}"] = date
                else: 
                    items["{{" + ser[x] + "}}"] = values[x]
            if 'percent' in our_type[x]:
                format_variable = re.search('[0-9]', our_type[x]).group(0)
                new_value = values[x] * 100
                if  isinstance(new_value, str):
                    if format_variable == 0:
                        y = str("{:.0f}%".format(values[x] * 100))
                        val = str(y)
                        items["{{" + ser[x] + "}}"] = val
                    if format_variable == 1:
                        y = str("{:.1f}%".format(values[x] * 100))
                        val = str(y)
                        items["{{" + ser[x] + "}}"] = val 
                    if format_variable == 2:
                        y = str("{:.2f}%".format(values[x] * 100))
                        val = str(y)
                        items["{{" + ser[x] + "}}"] = val
                    if format_variable == 3:
                        y = str("{:.3f}%".format(values[x] * 100))
                        val = str(y)
                        items["{{" + ser[x] + "}}"] = val
                elif isinstance(new_value, float):
                    val = f'{new_value:,.{format_variable}f}%'
                    val = str(val)
                    items["{{" + ser[x] + "}}"] = val
            if 'money' in our_type[x]:
                format_variable = re.search('[0-9]', our_type[x]).group(0)
                val = f'${values[x]:,.{format_variable}f}'
                val = str(val)
                items["{{" + ser[x] + "}}"] = val

    for value in items:
        print(value, items[value])
    print(type(items["{{submit.date}}"]))
    return items

#______________________end of GET DATA method_____________________________________________


def jsonDUMPER(items):
    jsonDUMP = []
    for key in items:
        jsonDUMP.append({'replaceAllText': {'containsText': {'text': key,'matchCase': True},'replaceText': items[key]}})

    #print(jsonDUMP)
    return jsonDUMP


#______Main Method sends data up to the Google Drive server
def main(items, jsonDUMP):
#Get path for exectuable or script
    if getattr(sys, 'frozen', False):
        print("EXECUTION TYPE: EXECUTABLE")
        application_path_excel = os.path.dirname(sys.executable) + "/"
        application_path_json_folder = os.path.dirname(sys.executable) + "/folder_data.json"
        application_path_json= os.path.dirname(sys.executable) + "/credentials.json"
        directory_list = os.scandir(application_path_excel)
        for x in directory_list:
            if '.xlsx' in x.name:
                excel_document = application_path_excel + x.name 
                excel_name = x.name
    elif __file__:
        print("EXECUTION TYPE: SCRIPT")
        application_path_json_folder = os.path.dirname(__file__) + "folder_data.json"
        application_path_json = os.path.dirname(__file__) + "credentials.json"
        application_path_excel = os.path.dirname(__file__) 
        directory_list = os.scandir()
        for x in directory_list:
            if '.xlsx' in x.name:
                excel_document = application_path_excel + x.name
                excel_name = x.name
#______________________________________________________________
#Check for credneitals
    json_document = json.loads(open(application_path_json_folder).read())
    template_fund3_id = json_document['drive_id']['master_template_fund3']
    master_folder_id = json_document['drive_id']['master_folder']
    portfolio_template_id = json_document['drive_id']['template_portfolio_comp']
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                application_path_json, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
#______________________________________________________________

#object representing the http request tether to  the slides API via credentials. Used to then call presentations
    service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)


#------------------------------------------------------
    #Call the Slides API // 
    '''presentation = service.presentations().get(
        presentationId='16nmnFPBu4cN8EzRBleeT7oNRbijAhH4cD0NAt2b3dFU').execute()
    slides = presentation.get('slides')'''

#______________________________________________________________
    #Create a copy of the presentation and portfolio companies, add date and Name
    today = date.today()
    copy_title =  'Fund 3 Deck: Updated-' + today.strftime("%b-%d-%Y") 
    body = {
        'name': copy_title
    }
    drive_response = drive_service.files().copy(
        fileId=template_fund3_id, body=body).execute()
    presentation_copy_id = drive_response.get('id')
#Portfolio comp:
    today = date.today()
    copy_title =  'Portfolio Companies: Updated ' + today.strftime("%b-%d-%Y") 
    body = {
        'name': copy_title
    }
    drive_response_port = drive_service.files().copy(
        fileId=portfolio_template_id, body=body).execute()
    portfolio_copy_id = drive_response_port.get('id')

    print("PASSING DATA UP TO THE SERVER")
#______________________________________________________________
 #Create an Archive Folder and Save the ID
    file_metadata = {
        'name': 'Archived Data: ' + today.strftime("%b-%d-%Y"),
        'mimeType': 'application/vnd.google-apps.folder',
        'parents' : [json_document["drive_id"]["archive_folder"]]
    }
    our_file = drive_service.files().create(body=file_metadata,
                                    fields='id').execute()
    folder_id = our_file.get('id')
#______________________________________________________________
 #upload Excel Document as .xlsx
    file_metadata = {'name': excel_name, 'parents': [folder_id]}
    media = MediaFileUpload(excel_document,
                        mimetype='image/jpeg')
    excel_file = drive_service.files().create(body=file_metadata,
                                    media_body=media,
                                    fields='id').execute()
    excel_id = excel_file.get('id')
    #os.remove(excel_document)
#______________________________________________________________
#Send request-------------------  
    request_num = 0
    body = {'requests': jsonDUMP}
    response = service.presentations().batchUpdate(
                presentationId=presentation_copy_id, body=body).execute()
    '''for key in items:
        try:
            requests = [
                {
                    'replaceAllText': {
                        'containsText': {
                            'text': key,
                            'matchCase': True
                        },
                        'replaceText': items[key]
                    }
                }
            ]
            body = {
            'requests': requests
            }
            response = service.presentations().batchUpdate(
                presentationId=presentation_copy_id, body=body).execute()
            time.sleep(1)
            request_num+=1
            print("HELLO")
            print(request_num)
        except:
            print("We have exceeded Google's quota for API requests. Going to hangout of 100 seconds")
            time.sleep(110) 
            #Rebuild API
            service = build('slides', 'v1', credentials=creds)
            drive_service = build('drive', 'v3', credentials=creds)
            response = service.presentations().batchUpdate(
                presentationId=presentation_copy_id, body=body).execute()
            request_num+=1
            print(request_num)
            portfolio_response = service.presentations().batchUpdate(
                presentationId=portfolio_copy_id, body=body).execute()'''

    #______________________________________________________________
    #move the copied and updated presentation into an archive folder


    file = drive_service.files().update(fileId=presentation_copy_id, 
                                addParents=folder_id,
                                removeParents=json_document["drive_id"]["master_folder"],
                                fields='id, parents').execute()
    file = drive_service.files().update(fileId=portfolio_copy_id, 
                                addParents=folder_id,
                                removeParents=json_document["drive_id"]["master_folder"],
                                fields='id, parents').execute()
#______________________________________________________________
    print("UPDATED")
if __name__ == '__main__':
    data = get_data()
    jsonDUMP = jsonDUMPER(data)
    main(data, jsonDUMP)