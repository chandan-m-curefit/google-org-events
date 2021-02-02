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

#Asia/Kolkata Timezone
ist = timezone(timedelta(hours=5,minutes=30))

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/admin.directory.user.readonly', 'https://www.googleapis.com/auth/calendar.readonly', 'https://www.googleapis.com/auth/admin.directory.orgunit.readonly']
SERVICE_ACCOUNT_FILE = 'service.json'
SERVICE_ACCOUNT_SUBJECT = 'samplecf@joefix.in'
#subject is the admin account with access rights

creds = None
serviceAdmin = None #Directory API
serviceCal = None #Calendar API

def dateFormat(dateStr,endOfDay=False):
    if dateStr==None:
        return None
    try:
        date = datetime.strptime(dateStr,'%d-%m-%Y').replace(tzinfo=ist)
        if endOfDay:
            date = date.replace(hour=23, minute=59, second=59)
        return date.isoformat()
    except ValueError as e:
        print("Error:",e)
        exit()

def connect_oauth():
    #Connect via oAuth - Authorize by browser login
    #Needs credentials.json for the GCP Project
    #Saves the credentials in a token.pickle file
    global serviceAdmin,serviceCal
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    serviceAdmin = build('admin', 'directory_v1', credentials=creds)
    serviceCal = build('calendar', 'v3', credentials=creds)
    
def connect_service():
    #Connects via service account credentials file.
    #No need of manually authorizing.
    global serviceAdmin,serviceCal
    creds = None
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES, subject=SERVICE_ACCOUNT_SUBJECT)
    #subject is the admin account with access rights
    #Client ID must be authorized with the Required Scopes in Domain-wide Delegation
    creds = credentials#.with_subject('samplecf@joefix.in')

    serviceAdmin = build('admin', 'directory_v1', credentials=creds)
    serviceCal = build('calendar', 'v3', credentials=creds)

'''def connect_api_key(apiKey):
    serviceAdmin = None #Directory API
    serviceCal = None #Calendar API
    return serviceAdmin, serviceCal'''


#userStatus = ('active','suspended','deleted')
def listUsersHelper(userStatus=None,orgUnitPath=None, debug=False):
    userList = list()
    page_token = None
    query=""
    if orgUnitPath:
        query="orgUnitPath={} ".format(orgUnitPath)
    showDeleted=False
    if 'active'==userStatus: query+="isSuspended=False"
    elif 'suspended'==userStatus: query+="isSuspended=True"
    elif 'deleted'==userStatus: showDeleted="True"
    try:
        while True:
            results = serviceAdmin.users().list(customer='my_customer', maxResults=100, orderBy='givenName', query=query, showDeleted=showDeleted, pageToken=page_token).execute()
            users = results.get('users', [])
            for user in users:
                temp = {}
                temp['email'] = user['primaryEmail']
                temp['name'] = user['name']['fullName']
                temp['status'] = userStatus
                userList.append(temp)
            page_token = results.get('nextPageToken')
            if not page_token:
                break
    except HttpError as e:
        if debug: print("ERROR: Invalid Input!\n",e)
    
    return userList

def getUsers(active=False, suspended=False, deleted=False):
    if not serviceAdmin:
        print("ERROR: Not connected!")
        pass
    if not(active or suspended or deleted):
        active=True;suspended=True;deleted=True
    userList = list()
    #Get all Active Users
    if active:
        userList.extend(listUsersHelper('active',debug=True))

    #Get all Suspended Users
    if suspended:
        userList.extend(listUsersHelper('suspended'))

    #Get all Deleted Users
    if deleted:
        userList.extend(listUsersHelper('deleted'))
    
    return userList

#Date format - DD-MM-YYYY        
def getMeetingsForUser(user,startDate=None,endDate=None,showDeleted=False,debug=False):
    if not serviceCal:
        print("ERROR: Not connected!")
        pass
    startDate = dateFormat(startDate)
    endDate = dateFormat(endDate,endOfDay=True)
    page_token = None
    meetings = list()
    try:
        while True:
            results = serviceCal.events().list(calendarId=user, pageToken=page_token, singleEvents=True, orderBy='startTime', timeMin=startDate, timeMax=endDate, showDeleted=showDeleted).execute()
            events = results.get('items', [])
            for event in events:
                #print(event['summary'],event['start']['dateTime'])
                temp = {}
                temp['id'] = event['id']
                temp['summary'] = event['summary']
                temp['startTime'] = event['start']['dateTime']
                temp['endTime'] = event['end']['dateTime']
                temp['status'] = event['status']
                temp['attendees'] = event['attendees']
                meetings.append(temp)
            page_token = results.get('nextPageToken')
            if not page_token:
                break
    except HttpError as e:
        if debug: print("ERROR: User with email '"+email+"' not found!")

    return meetings

def getOrgUnits(orgUnitPath=None):
    results = serviceAdmin.orgunits().list(customerId='my_customer',orgUnitPath=orgUnitPath).execute()
    print(results)
    orgUnits = results.get("organizationUnits", [])
    for item in orgUnits:
        print(item['orgUnitId'], item['name'], item['orgUnitPath'])

def getUsersInOrgUnit(orgId=None, orgPath=None, debug=True):
    try:
        if orgId:
            results = serviceAdmin.orgunits().get(customerId='my_customer', orgUnitPath=orgId).execute()
            orgPath = results.get('orgUnitPath', None)

        return listUsersHelper(orgUnitPath=orgPath, debug=debug)
    except HttpError as e:
        if debug: print("ERROR: Invalid Input - orgId or orgPath Wrong!\n",e,"\n")
        return list()

def getDistinctMeetingsForOrg(startDate=None, endDate=None, orgId=None, orgPath=None, showDeleted=False):
    userList = getUsersInOrgUnit(orgId=orgId, orgPath=orgPath)
    meetSet = set()
    meetList = list()
    for user in userList:
        response = getMeetingsForUser(user=user['email'], startDate=startDate, endDate=endDate, showDeleted=showDeleted)
        for item in response:
            if item['id'] not in meetSet:
                meetSet.add(item['id'])
                meetList.append(item)
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
    
    
    #-------------------------------
    #1. Fetching list of Users - getUser
    if choice=="1":
        print("\n1. Fetching list of Users\n")
        #
        userList= getUsers() #Fetches all users
        #
        #userList = getUsers(active = True) #Fetches active users
        #userList = getUsers(suspended = True) #Fetches suspended users
        #userList = getUsers(deleted = True) #Fetches deleted users
        if userList:
            with open( 'userList.txt', 'w' ) as file:
                for item in userList:
                    file.write( json.dumps(item)+"\n" )
            print("\nCheck userList.txt for the result")
        else:
            print("\nNo User Found!")
    
    #-------------------------------
    #2. Fetch a list of meetings for 1 user.
    elif(choice=="2"):
        print("\n2. Fetch a list of meetings for 1 user\n")
        email = input("Enter Email ID - ")
        startDate = input("Enter Start Date (DD-MM-YYYY) - ")
        endDate = input("Enter End Date (DD-MM-YYYY) - ")
        if startDate=="":startDate = None
        if endDate=="":endDate = None
        #
        meetList = getMeetingsForUser(email, startDate, endDate,debug=True)
        #
        if meetList:
            with open( 'userMeetings.txt', 'w' ) as file:
                for item in meetList:
                    file.write( json.dumps(item)+"\n" )
            print("\nCheck userMeetings.txt for the result")
        else:
            print("\nNo Meetings Found!")

    #-------------------------------
    #3. Fetch list of all meetings for all Users
    elif(choice=="3"):
        print("\n3. Fetch list of all meetings for all Users\n")
        startDate = input("Enter Start Date (DD-MM-YYYY) - ")
        endDate = input("Enter End Date (DD-MM-YYYY) - ")
        if startDate=="":startDate = None
        if endDate=="":endDate = None
        
        userList= getUsers() #Fetches all users
        if userList:
            with open( 'allUserMeetings.txt', 'w' ) as file:
                for user in userList:
                    temp = user.copy()
                    meetList = getMeetingsForUser(user['email'], startDate, endDate)
                    temp['meetings'] = meetList
                    file.write( json.dumps(temp)+"\n")
            print("\nCheck allUserMeetings.txt for the result")
        else:
            print("\nNo User Found!")
        
    
    #-------------------------------
    #4. Fetch list of users in orgUnit
    elif choice=="4":
        print("\n4. Fetching list of Users in orgUnit\n")
        orgId = input("Enter OrgId - ")
        orgPath = input("Enter OrgPth - ")
        if orgId=="":orgId = None
        if orgPath=="":orgPath = None
        #
        userList= getUsersInOrgUnit(orgId=orgId, orgPath=orgPath) #Fetches all users in given orgId
        #
        if userList:
            with open( 'orgUserList.txt', 'w' ) as file:
                for item in userList:
                    file.write( json.dumps(item)+"\n" )
            print("\nCheck orgUserList.txt for the result")
        else:
            print("\nNo User Found!")
    
    #-------------------------------
    #5. Fetch distinct meetings in the orgUnit.
    elif(choice=="5"):
        print("\n5. Fetch distinct meetings in the orgUnit.\n")
        orgId = input("Enter OrgId - ")
        orgPath = input("Enter OrgPth - ")
        if orgId=="":orgId = None
        if orgPath=="":orgPath = None
        startDate = input("Enter Start Date (DD-MM-YYYY) - ")
        endDate = input("Enter End Date (DD-MM-YYYY) - ")
        if startDate=="":startDate = None
        if endDate=="":endDate = None
        #
        meetList = getDistinctMeetingsForOrg(startDate=startDate, endDate=endDate, orgId=orgId, orgPath=orgPath)
        #
        if meetList:
            with open( 'orgDistinctMeetings.txt', 'w' ) as file:
                for item in meetList:
                    file.write( json.dumps(item)+"\n" )
            print("\nCheck orgDistinctMeetings.txt for the result")
        else:
            print("\nNo Meetings Found!")


if __name__ == '__main__':
    #Connect to Google API
    connect_service() #Connect Using Service Account
    #connect_oauth() #Connect via oAuth - Authorize by browser login
    menuProgram()
    #orgList = getOrgUnits()
    #userList = getUsersInOrgUnit(orgId='id:03ph8a2z2jxbw6q')
    #meetList = getMeetingsForUser("sample2@joefix.in", "31-01-2021", "05-02-2021",debug=True)
    #meetList = getDistinctMeetingsForOrg( startDate="31-01-2021", endDate="05-02-2021",orgPath="/Child1")