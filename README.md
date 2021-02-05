# Python Sript to Fetch Users and their Meetings details.

## Functions:

1. getUsers(active=False, suspended=False, deleted=False)
    - Gets a list of users in the Organization
    - If no parameters are set, it fetches all types of users.
    - If any parameter is set, then it fetches that type of user.
            Response:
            [
                { "email": "", "name": "", "status": ""},
                ....
            ]

2. getMeetingsForUser(user,startDate=None,endDate=None,debug=False)
    - Gets a list of meetings for the specified user
    - Date format: DD-MM-YYYY
    - If dates are not specified, then all meetings are fetched
            Response:
            [
                {"id": "", "summary": "", "startTime": "", "endTime": "", "status": "",  "attendees": [{"email": "", "responseStatus": ""},.... ]},
                ....
            ]

3. getDistinctMeetingsForOrg(startDate=None, endDate=None, orgId=None, orgPath=None)
    - Gets  list of distinct meetings in the specified organizational unit.
    - Can either provide orgId or orgPath
    - To fetch details for the entire company, leave orgId and orgPath blank.
    - If company is divided in to different orgUnits (ex. for each team)
        - Then the organization structure can be found at https://admin.google.com/ >  Organizational units
    - Steps:
        - Fetches a list of users in the orgUnit
        - Then Fetches meetings for each user
        - Then keep distinct meetings
                Response:
                [
                    {"id": "", "summary": "", "startTime": "", "endTime": "", "status": "", "attendees": [{"email": "", "responseStatus": ""},.... ]},
                    ....
                ]
                
### Google Workspace Admin SDK > Directory API

## Steps to Run:

1. Enable DirectoryAPI - Crete Cloud Platform Project - From the admin account
    https://developers.google.com/admin-sdk/directory/v1/quickstart/python
    - Click on Enable Directory API
    - Create a new GCP Project
    - Download credentials.json and place in working directory

2. Enable CalendrAPI in the GCP project
    https://console.cloud.google.com/apis/library/calendar-json.googleapis.com
    - Select the Project
    - Enable Calendar API.

3. Create a Service Account:
    https://console.cloud.google.com/iam-admin/serviceaccounts
    - Select the Project
    - Click on Create service account
        - Enter name,email,description.
        - Skip optional parts
    - In service account page, 
        - Click on View Client ID &  Copy the Client ID
        - Click on Actions > Create key > Select JSON > Create
        - Move the downloded .json file to Working Directory and save as service.json

4. Add Domain-wide Delegation for the created Service account
    - From https://admin.google.com/
    - Go to Security > API controls > Manage Domain Wide Delegation
    - Click on Add new
        - Paste the copied Client ID, and add the following Scopes:
            - https://www.googleapis.com/auth/admin.directory.user.readonly
            - https://www.googleapis.com/auth/calendar.readonly
            - https://www.googleapis.com/auth/admin.directory.orgunit.readonly

5. Install Google Client Library
        pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
6. Change the following attribute in user_meetings.py:
        - SERVICE_ACCOUNT_SUBJECT = "admin-user-email"
        
        Make sure
        - OAUTH_CREDENTIALS_FILE = 'credentials.json'
        - SERVICE_ACCOUNT_FILE = 'service.json'
    OR
        - You can pass these as command line arguments 
7. Run the python script:
    - `python user_meetings.py`
    OR
    - `python user_meetings.py "path-to-credentials-file" "admin-account-email"`


On the command line you will have these options:

    1. Fetch list of users
    2. Fetch list of all meetings for 1 User
    3. Fetch list of all meetings for all Users
    4. Fetch list of users in orgUnit
    5. Fetch distinct meetings in the orgUnit
    
Then enter details like `email-id, start-date, end-date`.
(If start-date and end-date are left blank, then there is no bounds on date)

### Authentication:

- Use Service Account
    - `SERVICE_ACCOUNT_SUBJECT = "admin-user-email"`
    - `SERVICE_ACCOUNT_FILE = 'service.json'`
    - `python user_meetings.py "path-to-credentials-file" "admin-account-email"`
- Use O-Auth
    - `OAUTH_CREDENTIALS_FILE = 'credentials.json'`
    - `python user_meetings.py`
    - Then login to the google account on the browser.

    