###
# Proof of concept script to parse our running club's training schedule, and produce an Alexa Flash Briefing Skill
### TODO:
# Schedule execution
# Add sftp upload to club website
# publish skill formally so it's available to all
#
### 
# Original Idea: Kevin Queenan
# Author: Richard Leyton (richard@leyton.org), @rleyton
###
# Script credits: The core Google Sheets logic is based on quickStart.py from https://developers.google.com/sheets/api/quickstart/python
###

from __future__ import print_function
import httplib2
import os
import json
import datetime
import dateparser
from uuid import uuid4
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_CONFIG = 'appconfig.json'
APPLICATION_NAME = 'Bellahouston Road Runners training schedule - Alexa Flash Briefing generator'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    with open(APPLICATION_CONFIG,'r') as data:
            config=json.load(data)

    if config['spreadsheetId'] and config['rangeName']:
        output=[]
        result = service.spreadsheets().values().get(
            spreadsheetId=config['spreadsheetId'], range=config['rangeName']).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
        else:
            if config['columns']:
                now=datetime.datetime.now()
                enddate=dateparser.parse("one week from now")
                for row in values:
                    if len(row)>3:
                        session={}
                        session['uid'] = "urn:uuid:" + str(uuid4())
                        sessiondate=dateparser.parse(row[0])

                        # we don't care about past sessions, or those too far off
                        if sessiondate is None or (sessiondate<now or sessiondate>enddate):
                            continue

                        dayofsession=sessiondate.strftime("%A")
                        session['updateDate']=sessiondate.isoformat()+"Z"
                        #datetime.datetime.utcnow().isoformat() + "Z"
                        session['titleText']=config['titleText']
                        session['redirectUrl']=config['redirectUrl']
                        # Print columns A and E, which correspond to indices 0 and 4.
                        session['mainText']='On %s, %s will run %s' % (dayofsession,row[3],row[4])
                        output.append(session)

        print(json.dumps(output,indent=4))
                #print(config['columns'])

            #for row in values:
                # Print columns A and E, which correspond to indices 0 and 4.
             #   print('%s, %s' % (row[0], row[4]))


if __name__ == '__main__':
    main()

