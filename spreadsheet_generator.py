"""
Shows basic usage of the Sheets API. Prints values from a Google Spreadsheet.
"""
from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from pprint import pprint
from googleapiclient import discovery

import argparse
import mutagen
import os
import sys

def write_sheets_api(all_metadata):
    # Setup the Sheets API
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    store = file.Storage('credentials.json')
    creds = store.get()
    os.remove('credentials.json')
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    
    # Call the Sheets API
    spreadsheet_id = '1ZyxQIZAJ49wyJzYsobnUJzfRaIDgyFmZAtl3yDLAccM'
    result = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, 
                                                    range='All songs', 
                                                    valueInputOption='RAW', 
                                                    body=all_metadata)
    response = result.execute()
    print(response)
    return response

def read_file_metadata(filename):
    return mutagen.File(filename)
    


if __name__ == '__main__':
    SPREADSHEET_ID = '1ZyxQIZAJ49wyJzYsobnUJzfRaIDgyFmZAtl3yDLAccM'
    parser = argparse.ArgumentParser(description='Populate Google Spreadsheet \
                                     with data from a music library stored in \
				     a directory on a local filesystem.')
    parser.add_argument('-d', '--directory', nargs=1, default="C:\\Users\\Renata\\Music", metavar='directory')
    args = parser.parse_args()

    directory = args.directory[0] if isinstance(args.directory, list) else args.directory
    if not os.path.isdir(directory) or not os.path.exists(directory):
        print('directory is not a directory or does not exist. Specify an existing directory.')
        sys.exit(404)

    values = []
    for filename in os.listdir(directory):
        full_filename = os.path.join(directory, filename)
        if os.path.isfile(full_filename):
            new_row = []
            metadata = read_file_metadata(full_filename)
            CELLS = ['title', 'albumartist', 'album', 'date', 'genre', 
                     'tracknumber', 'tracktotal', 'comment']
            for cell in CELLS:
                try:
                    new_row.append(metadata[cell][0])
                except:
                    new_row.append('')
            filetype = full_filename.split('.')[-1]
            new_row.insert(1, filetype)
            values.append(new_row)

    body = { 'values': values }
    response = write_sheets_api(body)
