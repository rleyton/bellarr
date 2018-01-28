###
# Proof of concept script to parse our running club's training schedule, and produce an Alexa Flash Briefing Skill
###
# Copyright (c) 2018 Bellahouston Road Runners
###
# Original Idea: Kevin Queenan
# Author: Richard Leyton (richard@leyton.org), @rleyton
###
# Script credits: The core Google Sheets logic is based on quickStart.py from https://developers.google.com/sheets/api/quickstart/python
###
# TODO:
# Schedule execution
# Add sftp upload to club website
# publish skill formally so it's available to all members
#
###

from __future__ import print_function
import httplib2
import os
import json
import datetime
import dateparser
import ssl
import re
import pdfkit
import feedparser
import requests
import pytz
import tempfile
import os
from icalendar import Calendar, Event
from ftplib import FTP_TLS,FTP
from yattag import Doc
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

DEBUG=1


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

def fetchNews():
    options = {
        'page-size': 'A4',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
    }
    css='test.css'
    pdfkit.from_file('test.html', options=options,output_path="test.pdf")
    url="https://www.bellahoustonroadrunners.co.uk/feed/"
    feed=requests.get(url)
    news=feedparser.parse(feed)
    True

def process_sheet(service, config, configblock):
    result = service.spreadsheets().values().get(
        spreadsheetId=config['spreadsheetId'], range=config['rangeName']).execute()
    values = result.get('values', [])
    output=[]
    if not values:
        print('No data found.')
    else:
        if config['columns']:
            # pick out the dates we care about
            now = datetime.datetime.now()
            if 'daterange' in config:
                enddate = dateparser.parse(config['daterange'])
                if DEBUG:
                    print("Forcing enddate "+config['daterange'])
            else:
                enddate = dateparser.parse("one week from now")

            for row in values:
                if len(row) > 3:
                    if DEBUG:
                        print("Processing "+row[0])

                    session = {}
                    session['uid'] = "urn:uuid:" + str(uuid4())

                    # weekend training needs special handling
                    if configblock=='training' and re.search('\/',row[0]):
                        # just the first date will do
                        sessiondate = dateparser.parse(row[0].split('/')[0])
                    else:
                        sessiondate = dateparser.parse(row[0])

                    # we don't care about past sessions, or those too far off
                    if sessiondate is None or (sessiondate < now or sessiondate > enddate):
                        continue

                    dayofsession = sessiondate.strftime("%A")
                    dayandmonth = sessiondate.strftime("%a %b %-d")

                    session['updateDate'] = sessiondate.isoformat() + "Z"
                    # datetime.datetime.utcnow().isoformat() + "Z"
                    session['titleText'] = config['titleText']
                    session['redirectUrl'] = config['redirectUrl']
                    # Print columns A and E, which correspond to indices 0 and 4.
                    if 'formatRule' in config:
                        session['mainText']=eval(config['formatRule'])
                    else:
                        # special case
                        if configblock=='training' and dayofsession=='Saturday':
                            session['mainText'] = 'At the weekend, target %s' % (row[4])
                        else:
                            session['mainText'] = 'On %s, %s will run %s' % (dayofsession, row[3], row[4])
                    output.append(session)
    if len(output)==0 and DEBUG==1:
        print("No data extracted for configblock")
    return output

def process_output(output,masterconfig,config):
    now = datetime.datetime.now()
    ftp = None
    if masterconfig['upload']:
        ftp = FTP(masterconfig['upload']['host'])
        ftp.login(masterconfig['upload']['user'], masterconfig['upload']['password'])

    if output and 'alexa' in config and config['alexa']!=0:
        # just now we only want training
        with open("bellatraining.json", 'w') as f:
            print(json.dumps(output['training']+output['duty'],indent=4),file=f)

        if ftp is not None:
            with open("bellatraining.json",'rb') as r:
                ftp.storlines('STOR bellatraining.json',r)

    if output and 'ical' in config and config['ical']!=0:
        for item in output:
            # just a training/duty roster atm; could make race calendar
            #if item not in ['training','duty']:
            #    continue

            cal = Calendar()
            cal.add('prodid', '-//My calendar product//mxm.dk//')
            cal.add('version', '2.0')

            if DEBUG:
                print("ical processing "+item)
            for element in output[item]:
                event=Event()
                event.add('summary',element['titleText'])
                event.add('description',element['mainText'])
                event.add('dtstart',dateparser.parse(element['updateDate']) + datetime.timedelta(hours=18,minutes=30))
                event.add('dtend', dateparser.parse(element['updateDate']) + datetime.timedelta(hours=20,minutes=00))
                cal.add_component(event)

            f=open("bellacalendar-"+item+".ics","wb")
            f.write(cal.to_ical())
            f.close()
            ftp.storlines('STOR bellacalendar-'+item+".ics",open("bellacalendar-"+item+".ics",'rb'))

    if output and 'html' in config and config['html']!=0:
        doc,tag,text,line = Doc().ttl()
        doc.asis('<!DOCTYPE html>')
        with tag('html',lang='en'):
            with tag('head'):
                doc.asis('<meta charset="utf-8">')
                doc.asis('<meta name="viewport" content="width=device-width, initial-scale=1">')
                doc.asis('<link rel="stylesheet" href="thisweek.css">')
            with tag('body'):
                with tag('center'):
                    with tag('h1'):
                        text('Bellahouston Road Runners')
                    with tag('h2'):
                        text("Running diary from %s " % (now.strftime("%a %b %-d %Y")) )
                for item in output:
                    if DEBUG:
                        print("HTML processing "+item)
                    with tag('h2'):
                        text(masterconfig['ingest'][item]['titleText'])
                    for element in output[item]:
                        with tag('ul'):
                            line('li',element['mainText'])

        with open("thisweek.html",'w') as f:
            print(doc.getvalue(),file=f)
        pdfkit.from_file("thisweek.html","calendar.pdf")
        if ftp is not None:
            ftp.storlines('STOR thisweek.html',open("thisweek.html",'rb'))
            ftp.storlines('STOR thisweek.css', open("thisweek.css", 'rb'))
            ftp.storbinary('STOR calendar.pdf',open("calendar.pdf",'rb'))


def main():
    """Parse schedule
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    with open(APPLICATION_CONFIG,'r') as data:
            config=json.load(data)

    #fetchNews()

    output={}

    if config['ingest']:
        for ingestItem in config['ingest']:
            if 'active' in config['ingest'][ingestItem] and config['ingest'][ingestItem]['active']!=0:
                if DEBUG:
                    print("START Ingesting "+ingestItem)

                output[ingestItem]=process_sheet(service,config['ingest'][ingestItem],ingestItem)

                if DEBUG:
                    print("END Ingesting "+ingestItem)

    process_output(output,config,config['export'])

if __name__ == '__main__':
    main()

