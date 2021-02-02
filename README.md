Python Sript to Fetch Users and their Meeting details.

Functions:

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
        - Can either provide orgIdor or orgPath
        - Steps:
            - Fetches a list of users in the orgUnit
            - Then Fetches meetings for each user
            - Then keep distinct meetings
        Response:
            [
                {"id": "", "summary": "", "startTime": "", "endTime": "", "status": "", "attendees": [{"email": "", "responseStatus": ""},.... ]},
                ....
            ]


Google Workspace Admin SDK > Directory API
https://developers.google.com/admin-sdk/directory/v1/quickstart/python
   
Steps to Run:

    1.Enable DirectoryAPI - Crete Cloud Platform Project - From the admin account
    2.Download credentials.json and place in working directory
    3.Enable CalendrAPI in the project
    4.Install Google Client Library - pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
    5.Run the python script
    6.Login to the Google Account
        Or
      Use Service Account
        - Create a service_account for the poject
        - Create and Download access key and place in working directoy as "service.json"
        - set SERVICE_ACCOUNT_SUBJECT to email id of account with acess rights
        - The Client ID must be authorized with the Required Scopes in Domain-wide Delegation

    On the command line you will have these options:
        1. Fetch list of users
        2. Fetch list of all meetings for 1 User
        3. Fetch list of all meetings for all Users
        4. Fetch list of users in orgUnit
        5. Fetch distinct meetings in the orgUnit
    Then enter details like email-id, start-date, end-date.
    (If start-date and end-date are left blank, then there is no bounds on date)
