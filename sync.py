
from __future__ import print_function

import json

import httplib2
import os
import locale
import sys
import re
import time
import random
import datetime
import StringIO

try:
    import vobject
    import requests
    from apiclient import discovery
    from googleapiclient.errors import HttpError
    from oauth2client import client
    from oauth2client import tools
    from oauth2client.file import Storage
    from dateutil.tz import tzlocal
except ImportError as e:
    print ("ERROR: Missing module - %s" % e.args[0])
    sys.exit(1)

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar' #.readonly'
CLIENT_SECRET_FILE = 'google_auth.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


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
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    # credentials = None
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES, redirect_uri='urn:ietf:wg:oauth:2.0:oob')
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def stringFromUnicode(string):
    return string.encode(locale.getlocale()[1] or
                         locale.getpreferredencoding(False) or
                         "UTF-8", "replace")

def PrintErrMsg(msg):
    PrintMsg(msg)


def PrintMsg(msg):
    if isinstance(msg, unicode):
        msg = stringFromUnicode(msg)
    sys.stdout.write(msg)


def DebugPrint(msg):
    # return
    PrintMsg(msg)


def ParseReminder(rem):
    matchObj = re.match(r'^(\d+)([wdhm]?)(?:\s+(popup|email|sms))?$', rem)
    if not matchObj:
        PrintErrMsg('Invalid reminder: ' + rem + '\n')
        sys.exit(1)
    n = int(matchObj.group(1))
    t = matchObj.group(2)
    m = matchObj.group(3)
    if t == 'w':
        n = n * 7 * 24 * 60
    elif t == 'd':
        n = n * 24 * 60
    elif t == 'h':
        n = n * 60

    if not m:
        m = 'popup'

    return n, m

def _LocalizeDateTime(dt):
    if not hasattr(dt, 'tzinfo'):
        return dt
    if dt.tzinfo is None:
        return dt.replace(tzinfo=tzlocal())
    else:
        return dt.astimezone(tzlocal())


def _RetryWithBackoff(method):
    for n in range(0, 10):
        try:
            return method.execute()
        except HttpError, e:
            error = json.loads(e.content)
            if error.get('code') == '403' and \
                            error.get('errors')[0].get('reason') \
                            in ['rateLimitExceeded', 'userRateLimitExceeded']:
                time.sleep((2 ** n) + random.random())
            else:
                raise


def ImportICS(batch, service, gcalendar, verbose=False, dump=False, reminder=None,
              ics=None, icsIsFile=True):

    def CreateEventFromVOBJ(ve):

        event = {}

        if verbose:
            print ("+----------------+")
            print ("| Calendar Event |")
            print ("+----------------+")

        if hasattr(ve, 'summary'):
            DebugPrint("SUMMARY: %s\n" % ve.summary.value)
            if verbose:
                print ("Event........%s" % ve.summary.value)
            event['summary'] = ve.summary.value

        if hasattr(ve, 'location'):
            DebugPrint("LOCATION: %s\n" % ve.location.value)
            if verbose:
                print ("Location.....%s" % ve.location.value)
            event['location'] = ve.location.value

        if not hasattr(ve, 'dtstart') or not hasattr(ve, 'dtend'):
            PrintErrMsg("Error: event does not have a dtstart and "
                        "dtend!\n")
            return None

        if ve.dtstart.value:
            DebugPrint("DTSTART: %s\n" % ve.dtstart.value.isoformat())
        if ve.dtend.value:
            DebugPrint("DTEND: %s\n" % ve.dtend.value.isoformat())
        if verbose:
            if ve.dtstart.value:
                print ("Start........%s" % \
                    ve.dtstart.value.isoformat())
            if ve.dtend.value:
                print ("End..........%s" % \
                    ve.dtend.value.isoformat())
            if ve.dtstart.value:
                print ("Local Start..%s" % \
                    _LocalizeDateTime(ve.dtstart.value))
            if ve.dtend.value:
                print ("Local End....%s" % \
                    _LocalizeDateTime(ve.dtend.value))

        if hasattr(ve, 'rrule'):

            DebugPrint("RRULE: %s\n" % ve.rrule.value)
            if verbose:
                print ("Recurrence...%s" % ve.rrule.value)

            event['recurrence'] = ["RRULE:" + ve.rrule.value]

        if hasattr(ve, 'dtstart') and ve.dtstart.value:
            # XXX
            # Timezone madness! Note that we're using the timezone for the
            # calendar being added to. This is OK if the event is in the
            # same timezone. This needs to be changed to use the timezone
            # from the DTSTART and DTEND values. Problem is, for example,
            # the TZID might be "Pacific Standard Time" and Google expects
            # a timezone string like "America/Los_Angeles". Need to find
            # a way in python to convert to the more specific timezone
            # string.
            # XXX
            # print ve.dtstart.params['X-VOBJ-ORIGINAL-TZID'][0]
            # print self.cals[0]['timeZone']
            # print dir(ve.dtstart.value.tzinfo)
            # print vars(ve.dtstart.value.tzinfo)

            start = ve.dtstart.value.isoformat()
            if isinstance(ve.dtstart.value, datetime.datetime):
                event['start'] = {'dateTime': start,
                                  'timeZone': gcalendar['timeZone']}
            else:
                event['start'] = {'date': start}

            if reminder:
                event['reminders'] = {'useDefault': False,
                                      'overrides': []}
                for r in reminder:
                    n, m = ParseReminder(r)
                    event['reminders']['overrides'].append({'minutes': n,
                                                            'method': m})

            # Can only have an end if we have a start, but not the other
            # way around apparently...  If there is no end, use the start
            if hasattr(ve, 'dtend') and ve.dtend.value:
                end = ve.dtend.value.isoformat()
                if isinstance(ve.dtend.value, datetime.datetime):
                    event['end'] = {'dateTime': end,
                                    'timeZone': gcalendar['timeZone']}
                else:
                    event['end'] = {'date': end}

            else:
                event['end'] = event['start']

        if hasattr(ve, 'description') and ve.description.value.strip():
            descr = ve.description.value.strip()
            DebugPrint("DESCRIPTION: %s\n" % descr)
            if verbose:
                print ("Description:\n%s" % descr)
            event['description'] = descr

        if hasattr(ve, 'organizer'):
            DebugPrint("ORGANIZER: %s\n" % ve.organizer.value)

            if ve.organizer.value.startswith("MAILTO:"):
                email = ve.organizer.value[7:]
            else:
                email = ve.organizer.value
            if verbose:
                print ("organizer:\n %s" % email)
            event['organizer'] = {'displayName': ve.organizer.name,
                                  'email': email}

        if hasattr(ve, 'attendee_list'):
            DebugPrint("ATTENDEE_LIST : %s\n" % ve.attendee_list)
            if verbose:
                print ("attendees:")
            event['attendees'] = []
            for attendee in ve.attendee_list:
                if attendee.value.upper().startswith("MAILTO:"):
                    email = attendee.value[7:]
                else:
                    email = attendee.value
                if verbose:
                    print (" %s" % email)

                event['attendees'].append({'displayName': attendee.name,
                                           'email': email})

        return event

    if dump:
        verbose = True

    f = sys.stdin

    if ics and icsIsFile:
        try:
            f = file(ics)
        except Exception, e:
            PrintErrMsg("Error: " + str(e) + "!\n")
            sys.exit(1)
    elif ics and not icsIsFile:
        f = ics

    while True:

        try:
            v = vobject.readComponents(f).next()
        except StopIteration:
            break

        for ve in v.vevent_list:

            event = CreateEventFromVOBJ(ve)

            if not event:
                continue

            if dump:
                continue

            if not verbose:
                batch.add(service.events().insert(calendarId=gcalendar["id"], body=event))
                # newEvent = _RetryWithBackoff(batch)
                # hLink = self._ShortenURL(newEvent['htmlLink'])
                # hLink = newEvent['htmlLink']
                # PrintMsg('New event added: %s\n' % hLink)
                continue

            PrintMsg("\n[S]kip [i]mport [q]uit: ")
            val = raw_input()
            if not val or val.lower() == 's':
                continue
            if val.lower() == 'i':
                batch.add(service.events().insert(calendarId=gcalendar["id"], body=event))
                # newEvent = _RetryWithBackoff(batch)
                # hLink = self._ShortenURL(newEvent['htmlLink'])
                # hLink = newEvent['htmlLink']
                # PrintMsg('New event added: %s\n' % hLink)
            elif val.lower() == 'q':
                sys.exit(0)
            else:
                PrintErrMsg('Error: invalid input\n')
                sys.exit(1)
    _RetryWithBackoff(batch)


def callback(request_id, response, exception):
    if exception:
        # Handle error
        print (exception)
    else:
        print ("Permission Id: %s" % response.get('id'))



def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    allCals = []
    calList = _RetryWithBackoff(service.calendarList().list())
    while True:
        for cal in calList['items']:
            allCals.append(cal)
        pageToken = calList.get('nextPageToken')
        if pageToken:
            calList = _RetryWithBackoff(
                service.calendarList().list(pageToken=pageToken))
        else:
            break

    for cal in allCals:
        if cal["summary"] == "zimbra":
            zimbra_cal = cal
            _RetryWithBackoff(service.calendars().delete(calendarId=zimbra_cal["id"]))
            break
    calendar = {
        'summary': 'zimbra',
        'timeZone': 'Europe/Kiev'
    }
    zimbra_cal = service.calendars().insert(body=calendar).execute()

    with open("zimbra_creds.json") as fil:
        zimbra_creds = json.load(fil)
    zimbra_resp = requests.get(
        'https://zimbra.sequans.com/home/{}/calendar?fmt=ics&auth=ba'.format(zimbra_creds["client_id"]),
        auth=(zimbra_creds["client_id"], zimbra_creds["client_secret"]))

    batch = service.new_batch_http_request(callback=callback)

    ImportICS(batch, service, zimbra_cal, ics=StringIO.StringIO(zimbra_resp.text), icsIsFile=False)


if __name__ == '__main__':
    main()
