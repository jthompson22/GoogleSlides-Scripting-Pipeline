from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import date
import os
import pandas as pd
import Read_Data 


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

# The ID of a sample presentation.
PRESENTATION_ID = '16nmnFPBu4cN8EzRBleeT7oNRbijAhH4cD0NAt2b3dFU'
template_presentation_id = '16nmnFPBu4cN8EzRBleeT7oNRbijAhH4cD0NAt2b3dFU'

def main():
    
    '''creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    #if os.path.exists('token.pickle'):
       # with open('token.pickle', 'rb') as token:
            #creds = pickle.load(token)
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
    #Call the Slides API // might need to update this
    presentation = service.presentations().get(
        presentationId='16nmnFPBu4cN8EzRBleeT7oNRbijAhH4cD0NAt2b3dFU').execute()
    slides = presentation.get('slides')

    
#Create a copy of the presentation_add_date and Name
    today = date.today()
    copy_title =  today.strftime("%b-%d-%Y") + 'Fund3-Deck'
    body = {
        'name': copy_title
    }
    drive_response = drive_service.files().copy(
        fileId=template_presentation_id, body=body).execute()
    presentation_copy_id = drive_response.get('id')


'''#update the copied presentation

    
    sheet1, sheet2 = Read_Data.Get_Data()
    #sheet1 data
    check1, rnd1, add1 = [str("{:.2f}".format(x)) if '.' in str(x) else str(x) for x in sheet1['Fund 1']]
    check2, rnd2, add2 = [str("{:.2f}".format(x)) if '.' in str(x) else str(x) for x in sheet1['Fund 2']]
    check3, rnd3, add3 = [str("{:.2f}".format(x)) for x in sheet1['Fund 3']]
    #sheet2 data
    grossROIC, netTVPI, gradRate, netIRR = [str("{:.2f}".format(x)) for x in sheet2['A']]
    print(grossROIC, netIRR,  netTVPI, gradRate)
    print(check2, rnd2, add2)
    '''
    requests = [
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

    ]

    body = {
        'requests': requests
    }
    
    response = service.presentations().batchUpdate(
        presentationId=presentation_copy_id, body=body).execute()
    
   # num_replacements = 0
    #for reply in response.get('replies'):
     #   num_replacements += reply.get('replaceAllText') \
      #      .get('occurrencesChanged')
    #print('Created presentation for %s with ID: %s' %
     #   (customer_name, presentation_copy_id))
    #print('Replaced %d text instances' % num_replacements)
'''
if __name__ == '__main__':
    main()