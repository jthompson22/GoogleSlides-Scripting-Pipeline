from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient.http import MediaFileUpload
from datetime import date
import os
import pandas as pd
import Read_Data 


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

# The ID of a sample presentation.
master_folder_ID = '1zMrdfLBZdtuGpEq1oE_hGyRkwU9BXARw'
template_presentation_id = '16nmnFPBu4cN8EzRBleeT7oNRbijAhH4cD0NAt2b3dFU'
archive_folder_ID = '1AwJA-CqCUSFxbVy8pYaM8RxnZNW_05dP'

def main():
    
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

#object representing the authentication into the slides API via credentials. Used to then call presentations
    service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)


#------------------------------------------------------
    #Call the Slides API // 
    presentation = service.presentations().get(
        presentationId='16nmnFPBu4cN8EzRBleeT7oNRbijAhH4cD0NAt2b3dFU').execute()
    slides = presentation.get('slides')

    
    #Create a copy of the presentation_add_date and Name
    today = date.today()
    copy_title =  'Fund 3 Deck: Updated-' + today.strftime("%b-%d-%Y") 
    body = {
        'name': copy_title
    }
    drive_response = drive_service.files().copy(
        fileId=template_presentation_id, body=body).execute()
    presentation_copy_id = drive_response.get('id')


    '''folder_response = drive_service.files().list(q= "name contains 'Fund 3 Deck: Updated'",
                                            spaces='Master-Script',
                                            fields='nextPageToken, files(id, name)').execute()
    for file in folder_response.get('files', []):
    # Process change
        print('Found file: %s (%s)' % (file.get('name'), file.get('id')))'''


    #Create an Archive Folder and Save the ID
    file_metadata = {
        'name': 'Archived Data: ' + today.strftime("%b-%d-%Y"),
        'mimeType': 'application/vnd.google-apps.folder',
        'parents' : ['1AwJA-CqCUSFxbVy8pYaM8RxnZNW_05dP']
    }
    our_file = drive_service.files().create(body=file_metadata,
                                    fields='id').execute()
    folder_id = our_file.get('id')

    #upload Excel Document
    file_metadata = {'name': 'data.xlsx', 'parents': [folder_id]}
    media = MediaFileUpload('data.xlsx',
                        mimetype='image/jpeg')
    excel_file = drive_service.files().create(body=file_metadata,
                                    media_body=media,
                                    fields='id').execute()
    excel_id = excel_file.get('id')



#update the copied presentation


    
    '''items = Read_Data.Get_Data()
    #sheet1 data
    #check1, rnd1, add1 = [str("{:.2f}".format(x)) if '.' in str(x) else str(x) for x in sheet1['Fund 1']]
    #check2, rnd2, add2 = [str("{:.2f}".format(x)) if '.' in str(x) else str(x) for x in sheet1['Fund 2']]
    #check3, rnd3, add3 = [str("{:.2f}".format(x)) if '.' in str(x) else str(x) for x in sheet1['Fund 3']]
    #sheet2 data
    #grossROIC, netTVPI, gradRate, netIRR = [str("{:.2f}".format(x)) if '.' in str(x) else str(x) for x in sheet2['A']]
    #print(grossROIC, netIRR,  netTVPI, gradRate)
    #print(check2, rnd2, add2)
    
    for key in items:
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
            presentationId=presentation_copy_id, body=body).execute()'''

    '''requests = [
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{init-check-size-one}}',
                    'matchCase': True
                },
                'replaceText': check1
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{avg-rnd-size-one}}',
                    'matchCase': True
                },
                'replaceText': rnd1
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{to-add-rich-one}}',
                    'matchCase': True
                },
                'replaceText': add1
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{init-check-size-two}}',
                    'matchCase': True
                },
                'replaceText': check2
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{avg-rnd-size-two}}',
                    'matchCase': True
                },
                'replaceText': rnd2
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{to-add-rich-two}}',
                    'matchCase': True
                },
                'replaceText': add2
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{init-check-size-three}}',
                    'matchCase': True
                },
                'replaceText': check3
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{avg-rnd-size-three}}',
                    'matchCase': True
                },
                'replaceText': rnd3
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{to-add-rich-three}}',
                    'matchCase': True
                },
                'replaceText': add3
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{netIRR}}',
                    'matchCase': True
                },
                'replaceText': netIRR
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{gross-roic}}',
                    'matchCase': True
                },
                'replaceText': grossROIC
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{seedgrad}}',
                    'matchCase': True
                },
                'replaceText': gradRate
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{netTVPI}}',
                    'matchCase': True
                },
                'replaceText': netTVPI
            }
        }

    ]'''

    '''body = {
        'requests': requests
    }
    
    response = service.presentations().batchUpdate(
        presentationId=presentation_copy_id, body=body).execute()'''
    
   # num_replacements = 0
    #for reply in response.get('replies'):
     #   num_replacements += reply.get('replaceAllText') \
      #      .get('occurrencesChanged')
    #print('Created presentation for %s with ID: %s' %
     #   (customer_name, presentation_copy_id))
    #print('Replaced %d text instances' % num_replacements)

if __name__ == '__main__':
    main()