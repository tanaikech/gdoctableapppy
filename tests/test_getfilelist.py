#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from gdoctableapppy import gdoctableapp

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents']


def main():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    resource = {
        "oauth2": creds,
        "documentId": "###",
    }
    result = gdoctableapp.GetTables(resource)
    print(result)


if __name__ == '__main__':
    main()
