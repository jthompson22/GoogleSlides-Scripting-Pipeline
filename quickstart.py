from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import date

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

# The ID of a sample presentation.
PRESENTATION_ID = '1AHncIHrvlV_73pbjYkKGyFVspD0hNEfzRGlVt9Mf7Ws'
template_presentation_id = '1AHncIHrvlV_73pbjYkKGyFVspD0hNEfzRGlVt9Mf7Ws'

def main():
    """Shows basic usage of the Slides API.
    Prints the number of slides and elments in a sample presentation.
    """
    creds = None
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
        presentationId='1AHncIHrvlV_73pbjYkKGyFVspD0hNEfzRGlVt9Mf7Ws').execute()
    slides = presentation.get('slides')

    #print('The presentation contains {} slides:'.format(len(slides)))
    #for i, slide in enumerate(slides):
        #if i == 43:
           #print('- Slide #{} contains {} '.format(
                #i + 1, len(slide.get('pageElements'))))

    
#Create a copy of the presentation_add_date and Name
    today = date.today()
    copy_title =  today.strftime("%b-%d-%Y") + 'Fund3-Deck'
    body = {
        'name': copy_title
    }
    drive_response = drive_service.files().copy(
        fileId=template_presentation_id, body=body).execute()
    presentation_copy_id = drive_response.get('id')
#update the copied presentation

    customer_name = "Test1"
    case_description = "Test2"
    total_portfolio = "Test3"
    requests = [
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{init-check-size}}',
                    'matchCase': True
                },
                'replaceText': customer_name
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{avg-rnd-size}}',
                    'matchCase': True
                },
                'replaceText': case_description
            }
        },
        {
            'replaceAllText': {
                'containsText': {
                    'text': '{{to-add-rich}}',
                    'matchCase': True
                },
                'replaceText': total_portfolio
            }
        }
    ]

    body = {
        'requests': requests
    }
    
    response = service.presentations().batchUpdate(
        presentationId=presentation_copy_id, body=body).execute()
    
    num_replacements = 0
    for reply in response.get('replies'):
        num_replacements += reply.get('replaceAllText') \
            .get('occurrencesChanged')
    print('Created presentation for %s with ID: %s' %
        (customer_name, presentation_copy_id))
    print('Replaced %d text instances' % num_replacements)

if __name__ == '__main__':
    main()