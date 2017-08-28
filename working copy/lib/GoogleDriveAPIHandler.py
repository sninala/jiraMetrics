import os
import httplib2
import argparse
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from apiclient import discovery
from apiclient.http import MediaFileUpload


class GoogleDriveAPIHandler(object):

    def __init__(self, app_name, client_secret_file, credentials_directory):
        self.app_name = app_name
        self.client_secret_file = client_secret_file
        self.flags = self.initialize_flags()
        self.scopes = 'https://www.googleapis.com/auth/drive.file'
        self.credentials_directory = credentials_directory

    @staticmethod
    def initialize_flags():
        try:
            flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
        except ImportError:
            flags = None
        return flags

    def get_credentials(self):
        credential_dir = self.credentials_directory
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir, 'jiraMetrics-drive-credentials.json')
        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self.client_secret_file, self.scopes)
            flow.user_agent = self.app_name
            if self.flags:
                credentials = tools.run_flow(flow, store, self.flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def get_google_drive_api_obj(self):
        credentials = self.get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('drive', 'v3', http=http)
        return service

    def upload_file_to_google_drive_folder(self, local_file_path, remote_folder_id):
        service = self.get_google_drive_api_obj()
        file_name = os.path.basename(local_file_path)
        file_metadata = {
            'name': file_name,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            "parents": [remote_folder_id]
        }
        media = MediaFileUpload(local_file_path,
                                mimetype='application/vnd.google-apps.spreadsheet',
                                resumable=True)
        file_obj = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print('File ID: %s' % file_obj.get('id'))


'''
if __name__ == '__main__':
    google_api = GoogleDriveAPIHandler('jiraMetrics', '../client_secret.json')
    local_file = r'D:\jiraMetrics\working copy\output\From 2015-current - Combined 2017-08-21.xlsx'
    google_api.upload_file_to_google_drive_folder(local_file, '0B66p8j8YNzuMLVllM3dvZTRvYVE')
'''
