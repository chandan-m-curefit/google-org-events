from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from datetime import datetime, timezone, timedelta
import json
import sys
import time
from multiprocessing import Pool

# ________________________________
# Asia/Kolkata Timezone
ist = timezone(timedelta(hours=5, minutes=30))

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/admin.directory.user.readonly',
          'https://www.googleapis.com/auth/calendar.readonly',
          'https://www.googleapis.com/auth/admin.directory.orgunit.readonly']

OAUTH_CREDENTIALS_FILE = 'credentials.json'
SERVICE_ACCOUNT_FILE = 'service.json'
SERVICE_ACCOUNT_SUBJECT = 'samplecf@joefix.in'
BATCH_SIZE = 10
# Subject is the admin account with access rights
# Command Line Arguments:
#   python user_meetings.py <path to credentials file> <admin account email>
argLen = len(sys.argv)
if argLen >= 2:
    OAUTH_CREDENTIALS_FILE = sys.argv[1]
    SERVICE_ACCOUNT_FILE = sys.argv[1]
if argLen >= 3:
    SERVICE_ACCOUNT_SUBJECT = sys.argv[2]

serviceAdmin = None  # Directory API
serviceCal = None  # Calendar API


# ___________________________________

def dateFormat(dateStr, endOfDay=False):
    if dateStr is None:
        return None
    try:
        date = datetime.strptime(dateStr, '%d-%m-%Y').replace(tzinfo=ist)
        if endOfDay:
            date = date.replace(hour=23, minute=59, second=59)
        return date.isoformat()
    except ValueError as e:
        print("Error:", e)
        exit()


def connect_oauth():
    # Connect via oAuth - Authorize by browser login
    # Needs credentials.json for the GCP Project
    # Saves the credentials in a token.pickle file
    global serviceAdmin, serviceCal
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                OAUTH_CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    serviceAdmin = build('admin', 'directory_v1', credentials=creds)
    serviceCal = build('calendar', 'v3', credentials=creds)


def connect_service():
    # Connects via service account credentials file.
    # No need of manually authorizing.
    global serviceAdmin, serviceCal
    creds = None
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES, subject=SERVICE_ACCOUNT_SUBJECT)
    # subject is the admin account with access rights
    # Client ID must be authorized with the Required Scopes in Domain-wide Delegation
    creds = credentials  # .with_subject('samplecf@joefix.in')

    serviceAdmin = build('admin', 'directory_v1', credentials=creds)
    serviceCal = build('calendar', 'v3', credentials=creds)


def connect_service_cal():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES, subject=SERVICE_ACCOUNT_SUBJECT)
    creds = credentials  # .with_subject('samplecf@joefix.in')
    return build('calendar', 'v3', credentials=creds)


# userStatus = ('active','suspended','deleted')
def listUsersHelper(userStatus=None, orgUnitPath=None):
    userList = list()
    page_token = None
    query = ""
    if orgUnitPath:
        query = "orgUnitPath={} ".format(orgUnitPath)
    showDeleted = False
    if 'active' == userStatus:
        query += "isSuspended=False"
    elif 'suspended' == userStatus:
        query += "isSuspended=True"
    elif 'deleted' == userStatus:
        showDeleted = True
    try:
        # page = 1
        while True:
            results = serviceAdmin.users().list(customer='my_customer', orderBy='givenName', query=query,
                                                showDeleted=showDeleted, pageToken=page_token).execute()
            users = results.get('users', [])
            # print("Page {}: {} Users Fetched.".format(page, len(users)))
            # page += 1
            for user in users:
                temp = {
                    'email': user.get('primaryEmail', ''),
                    'name': user.get('name', {}).get('fullName', ''),
                    'status': userStatus
                }
                userList.append(temp)
            page_token = results.get('nextPageToken')
            if not page_token:
                break
    except HttpError as e:
        print("ERROR: Invalid Input!\n", e)
    print("{} Total Users Found".format(len(userList)))
    return userList


def getUsers(active=False, suspended=False, deleted=False):
    if not serviceAdmin:
        print("ERROR: Not connected!")
        pass
    if not (active or suspended or deleted):
        active = True
        suspended = True
        deleted = True
    userList = list()
    # Get all Active Users
    if active:
        userList.extend(listUsersHelper('active'))

    # Get all Suspended Users
    if suspended:
        userList.extend(listUsersHelper('suspended'))

    # Get all Deleted Users
    if deleted:
        userList.extend(listUsersHelper('deleted'))

    return userList


# Date format - DD-MM-YYYY
def getMeetingsForUser(user, startDate=None, endDate=None, showDeleted=False):
    serviceCal = globals()['serviceCal']
    if not serviceCal:
        serviceCal = connect_service_cal()
        if not serviceCal:
            print("ERROR: Not connected!")
            return

    startDate = dateFormat(startDate)
    endDate = dateFormat(endDate, endOfDay=True)
    page_token = None
    meetings = list()
    try:
        # page = 1
        while True:
            results = serviceCal.events().list(calendarId=user, maxResults=2500, pageToken=page_token,
                                               singleEvents=True,
                                               orderBy='startTime', timeMin=startDate, timeMax=endDate,
                                               showDeleted=showDeleted).execute()
            events = results.get('items', [])
            # print("Page {}: {} Meetings Fetched.".format(page, len(events)))
            # page += 1
            for event in events:
                # print(event['summary'],event['start']['dateTime'])
                temp = {
                    'id': event.get('id', ''),
                    'summary': event.get('summary', ''),
                    'startTime': event.get('start', {}).get('dateTime', ''),
                    'endTime': event.get('end', {}).get('dateTime', ''),
                    'status': event.get('status', ''),
                    'attendees': event.get('attendees', [])
                }
                meetings.append(temp)
            page_token = results.get('nextPageToken')
            if not page_token:
                break
    except HttpError as e:
        print("ERROR:", e)

    return meetings


def getOrgUnits(orgUnitPath=None):
    if not serviceAdmin:
        print("ERROR: Not connected!")
        pass
    results = serviceAdmin.orgunits().list(customerId='my_customer', orgUnitPath=orgUnitPath).execute()
    print(results)
    orgUnits = results.get("organizationUnits", [])
    for item in orgUnits:
        id = item.get('orgUnitId', '')
        name = item.get('name', '')
        path = item.get('orgUnitPath', '')
        print(id, name, path)


def getUsersInOrgUnit(orgId=None, orgPath=None):
    if not serviceAdmin:
        print("ERROR: Not connected!")
        pass
    try:
        if orgId:
            results = serviceAdmin.orgunits().get(customerId='my_customer', orgUnitPath=orgId).execute()
            orgPath = results.get('orgUnitPath', None)
        userList = list()
        userList.extend(listUsersHelper(userStatus='active', orgUnitPath=orgPath))
        userList.extend(listUsersHelper(userStatus='suspended', orgUnitPath=orgPath))
        print("{} total users found in orgUnit.".format(len(userList)))
        return userList
    except HttpError as e:
        print("ERROR: Invalid Input - orgId or orgPath Wrong!\n", e, "\n")
        return list()


def getDistinctMeetingsForOrg(startDate=None, endDate=None, orgId=None, orgPath=None, showDeleted=False):
    if not (serviceCal and serviceAdmin):
        print("ERROR: Not connected!")
        pass
    total_start_time = time.time()
    userList = getUsersInOrgUnit(orgId=orgId, orgPath=orgPath)
    print("\nNo. of users in the OrgUnit: {}\n".format(len(userList)))
    meetSet = set()
    meetList = list()
    for num, user in enumerate(userList):
        start_time = time.time()
        response = getMeetingsForUser(user=user['email'], startDate=startDate, endDate=endDate, showDeleted=showDeleted)
        print("{}. {} : {} meetings found.".format(num, user['email'], len(response)))
        for item in response:
            if item['id'] not in meetSet:
                meetSet.add(item['id'])
                meetList.append(item)
        print("--- Time: {} seconds ---".format(time.time() - start_time))
    print("\n Number of Distinct Meetings in the Org: {}".format(len(meetList)))
    print("--- Total Time: {} seconds ---".format(time.time() - total_start_time))
    return meetList


def getDistinctMeetingsForOrgParallel(startDate=None, endDate=None, orgId=None, orgPath=None, showDeleted=False,
                                      batchSize=1):
    if not (serviceCal and serviceAdmin):
        print("ERROR: Not connected!")
        return
    total_start_time = time.time()
    userList = getUsersInOrgUnit(orgId=orgId, orgPath=orgPath)
    userListLen = len(userList)
    print("\nNo. of users in the OrgUnit: {}\n".format(len(userList)))
    meetSet = set()
    meetList = list()
    num = 1
    for i in range(0, len(userList), batchSize):
        start_time = time.time()
        batchRange = range(i, min(i + batchSize, userListLen))
        batch = list()
        for j in batchRange:
            batch.append((userList[j]['email'], startDate, endDate, showDeleted))
        with Pool(batchSize) as pool:
            responseArr = pool.starmap(getMeetingsForUser, batch)
        for j, num in enumerate(batchRange):
            response = responseArr[j]
            print("{}. {} : {} meetings found.".format(num, userList[num]['email'], len(response)))
            for item in response:
                if item['id'] not in meetSet:
                    meetSet.add(item['id'])
                    meetList.append(item)
        print("--- Time: {} seconds ---".format(time.time() - start_time))
    print("\n Number of Distinct Meetings in the Org: {}".format(len(meetList)))
    print("--- Total Time: {} seconds ---".format(time.time() - total_start_time))
    return meetList


def menuProgram():
    print("Select Option")
    print("1. Fetch list of users")
    print("2. Fetch list of all meetings for 1 User")
    print("3. Fetch list of all meetings for all Users")
    print("4. Fetch list of users in orgUnit")
    print("5. Fetch distinct meetings in the orgUnit")
    startDate = None
    endDate = None

    choice = input("Enter Option - ")

    # -------------------------------
    # 1. Fetching list of Users - getUser
    if choice == "1":
        print("\n1. Fetching list of Users\n")
        #
        userList = getUsers()  # Fetches all users
        #
        # userList = getUsers(active = True) #Fetches active users
        # userList = getUsers(suspended = True) #Fetches suspended users
        # userList = getUsers(deleted = True) #Fetches deleted users
        if userList:
            with open('userList.txt', 'w') as file:
                for item in userList:
                    file.write(json.dumps(item) + "\n")
            print("\nCheck userList.txt for the result")
        else:
            print("\nNo User Found!")

    # -------------------------------
    # 2. Fetch a list of meetings for 1 user.
    elif choice == "2":
        print("\n2. Fetch a list of meetings for 1 user\n")
        email = input("Enter Email ID - ")
        startDate = input("Enter Start Date (DD-MM-YYYY) - ")
        endDate = input("Enter End Date (DD-MM-YYYY) - ")
        if startDate == "": startDate = None
        if endDate == "": endDate = None
        #
        meetList = getMeetingsForUser(email, startDate, endDate)
        #
        if meetList:
            with open('userMeetings.txt', 'w') as file:
                for item in meetList:
                    file.write(json.dumps(item) + "\n")
            print("Number of meetings found: {}".format(len(meetList)))
            print("\nCheck userMeetings.txt for the result")
        else:
            print("\nNo Meetings Found!")

    # -------------------------------
    # 3. Fetch list of all meetings for all Users
    elif choice == "3":
        print("\n3. Fetch list of all meetings for all Users\n")
        startDate = input("Enter Start Date (DD-MM-YYYY) - ")
        endDate = input("Enter End Date (DD-MM-YYYY) - ")
        if startDate == "": startDate = None
        if endDate == "": endDate = None
        batchSize = BATCH_SIZE
        total_start_time = time.time()
        userList = getUsers(active=True, suspended=True)  # Fetches all active and suspended users
        userListLen = len(userList)
        print("\nNo. of users : {}\n".format(len(userList)))

        if userList:
            with open('allUserMeetings.txt', 'w') as file:
                for i in range(0, len(userList), batchSize):
                    start_time = time.time()
                    batchRange = range(i, min(i + batchSize, userListLen))
                    batch = list()
                    for j in batchRange:
                        batch.append((userList[j]['email'], startDate, endDate))
                    with Pool(batchSize) as pool:
                        responseArr = pool.starmap(getMeetingsForUser, batch)
                    for j, num in enumerate(batchRange):
                        response = responseArr[j]
                        print("{}. {} : {} meetings found.".format(num, userList[num]['email'], len(response)))
                        temp = userList[num].copy()
                        temp['meetings'] = response
                        file.write(json.dumps(temp) + "\n")
                    print("--- Time: {} seconds ---".format(time.time() - start_time))
                """for user in userList:
                    # if(user['status']!='deleted'):
                    temp = user.copy()
                    meetList = getMeetingsForUser(user['email'], startDate, endDate)
                    temp['meetings'] = meetList
                    file.write(json.dumps(temp) + "\n")"""
            print("--- Total Time: {} seconds ---".format(time.time() - total_start_time))
            print("\nCheck allUserMeetings.txt for the result")
        else:
            print("\nNo User Found!")


    # -------------------------------
    # 4. Fetch list of users in orgUnit
    elif choice == "4":
        print("\n4. Fetching list of Users in orgUnit\n")
        orgId = input("Enter OrgId - ")
        orgPath = input("Enter OrgPath - ")
        if orgId == "": orgId = None
        if orgPath == "": orgPath = None
        #
        userList = getUsersInOrgUnit(orgId=orgId, orgPath=orgPath)  # Fetches all users in given orgId
        #
        if userList:
            with open('orgUserList.txt', 'w') as file:
                for item in userList:
                    file.write(json.dumps(item) + "\n")
            print("\nCheck orgUserList.txt for the result")
        else:
            print("\nNo User Found!")

    # -------------------------------
    # 5. Fetch distinct meetings in the orgUnit.
    elif choice == "5":
        print("\n5. Fetch distinct meetings in the orgUnit.\n")
        orgId = input("Enter OrgId - ")
        orgPath = input("Enter OrgPath - ")
        if orgId == "": orgId = None
        if orgPath == "": orgPath = None
        startDate = input("Enter Start Date (DD-MM-YYYY) - ")
        endDate = input("Enter End Date (DD-MM-YYYY) - ")
        if startDate == "": startDate = None
        if endDate == "": endDate = None
        #
        meetList = getDistinctMeetingsForOrgParallel(
            startDate=startDate, endDate=endDate, orgId=orgId,
            orgPath=orgPath, batchSize=BATCH_SIZE
        )
        # meetList = getDistinctMeetingsForOrg(startDate=startDate, endDate=endDate, orgId=orgId, orgPath=orgPath)

        #
        if meetList:
            with open('orgDistinctMeetings.txt', 'w') as file:
                for item in meetList:
                    file.write(json.dumps(item) + "\n")
            print("\nCheck orgDistinctMeetings.txt for the result")
        else:
            print("\nNo Meetings Found!")


if __name__ == '__main__':
    # Connect to Google API
    connect_service()  # Connect Using Service Account
    menuProgram()

    # ------
    # orgList = getOrgUnits()
    # userList = getUsersInOrgUnit(orgId='id:03ph8a2z2jxbw6q')
    # meetList = getMeetingsForUser("sample2@joefix.in", "31-01-2021", "05-02-2021",debug=True)
    # meetList = getDistinctMeetingsForOrg( startDate="31-01-2021", endDate="05-02-2021",orgPath="/Child1")
