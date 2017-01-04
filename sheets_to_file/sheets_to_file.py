#!/usr/bin/python3
""" sheets_to_file.py:
    Reads data from a google sheet and writes that data to a file on the local disk in comma
    separated variable (csv) format.
"""

# Import Required Libraries (Standard, Third Party, Local) ****************************************
from __future__ import print_function
import httplib2
import file_logger
import logging
import os
import sys

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


# Authorship Info *********************************************************************************
__author__ = "Christopher Maue"
__copyright__ = "Copyright 2016, The RPi-Home Project"
__credits__ = ["Christopher Maue"]
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Christopher Maue"
__email__ = "csmaue@gmail.com"
__status__ = "Development"


# Class Definitions *******************************************************************************
class GoogleSheetsToFile(object):
    """ Class and methods necessary to read rows from a google sheet via google's api' and
    write that info to a file on local disk in CSV (comma separated variable) format 
    """
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
        try:
            import argparse
            self.flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
        except ImportError:
            self.flags = None
        self.home_dir = str()
        self.credential_dir = str()
        self.store = str()
        self.credentials = str()
        self.path = str()
        self.CLIENT_SECRET_FILE = str()
        self.SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
        self.CLIENT_SECRET_FILE = 'client_secret.json'
        self.APPLICATION_NAME = 'Device Schedule via Google Sheets API'


    def get_credentials(self):
        """ Gets valid user credentials from storage. If nothing has been stored, or if the stored
        credentials are invalid, the OAuth2 flow is completed to obtain the new credentials. 
        Returns: Credentials, the obtained credential.
        """
        self.home_dir = os.path.expanduser('~')
        self.credential_dir = os.path.join(self.home_dir, '.credentials')
        if not os.path.exists(self.credential_dir):
            self.logger.debug("Creating directory: %s", self.credential_dir)
            os.makedirs(self.credential_dir)
        self.credential_path = os.path.join(
            self.credential_dir, 'sheets.googleapis.com-python-quickstart.json')
        self.logger.debug("Setting credential path to: %s", self.credential_path)
        self.store = Storage(self.credential_path)
        self.logger.debug("Setting store to: %s", self.store)
        self.credentials = self.store.get()
        self.logger.debug("Getting credentials from store")
        if not self.credentials or self.credentials.invalid:
            self.logger.debug("Credentials not in store")
            self.path = os.path.dirname(sys.argv[0])
            self.logger.debug("System path is: %s", self.path)
            self.CLIENT_SECRET_FILE = os.path.join(self.path, "client_secret.json")
            self.logger.debug("Looking for json file at: %s", self.CLIENT_SECRET_FILE)
            self.flow = client.flow_from_clientsecrets(self.CLIENT_SECRET_FILE, self.SCOPES)
            self.flow.user_agent = self.APPLICATION_NAME
            if self.flags:
                self.credentials = tools.run_flow(self.flow, self.store, self.flags)
            else: # Needed only for compatibility with Python 2.6
                self.credentials = tools.run(self.flow, self.store)
            self.logger.debug('Storing credentials to ' + self.credential_path)
        self.logger.debug("Returning credentials to main program")
        return self.credentials


    def read_data(self, sheet_id=None, sheet_range=None):
        """ Returns all data from a specific sheet using the google sheets API
        """
        self.credentials = self.get_credentials()
        self.http = self.credentials.authorize(httplib2.Http())
        self.discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
        self.service = discovery.build('sheets', 'v4',
                                       http=self.http,
                                       discoveryServiceUrl=self.discoveryUrl)
        # Set sheet name and range to read
        if sheet_id is not None:
            self.spreadsheetId = sheet_id
        else:
            self.spreadsheetId = '1LJpDC0wMv3eXQtJvHNav_Yty4PQcylthOxXig3_Bwu8'
        self.logger.debug("Using sheet id: %s", self.spreadsheetId)
        if sheet_range is not None:
            self.rangeName = sheet_range
        else:
            self.rangeName = "fylt1!A3:L"
        self.logger.debug("Reading data from range: %s", self.rangeName)
        # Read data from sheet/range specified
        self.result = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheetId, range=self.rangeName).execute()
        self.values = self.result.get('values', [])
        self.logger.debug("Read from table: %s", self.values)

        if not self.values:
            self.logger.debug("No data found.  Returning NONE to main")
            return None
        else:
            self.logger.debug("Returning data to main")
            return self.values


    def write_to_file(self, lines=None, output_file=None):
        self.file = open(output_file, "w")
        for i, j in enumerate(lines):
            if isinstance(j, list):
                for k, l in enumerate(j):
                    if k > 0:
                        self.file.write(", ")
                    self.file.write(l)
                self.file.write("\n")
            elif isinstance(j, str):
                self.file.write(j)
                self.file.write("\n")
        self.file.close()



if __name__ == "__main__":
    debug_file, info_file = file_logger.setup_log_files(__file__)
    logger = file_logger.setup_log_handlers(__file__, debug_file, info_file)
    sheet_object = GoogleSheetsToFile(logger=logger)
    
    results = sheet_object.read_data(
        sheet_id='1LJpDC0wMv3eXQtJvHNav_Yty4PQcylthOxXig3_Bwu8',
        sheet_range="fylt1!A3:L")
    print(results)
    output_file = "c:\python_logs\schedule.txt"
    sheet_object.write_to_file(lines=results, output_file=output_file)