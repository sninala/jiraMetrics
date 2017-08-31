import os
import sys
import time
import base64
import httplib2
import oauth2client
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from email.mime.text import MIMEText
from apiclient import errors


class EmailSender(object):
    def __init__(self, app_name, client_secret_file, credentials_directory):
        self.app_name = app_name
        self.client_secret_file = client_secret_file
        self.flags = self.initialize_flags()
        self.scopes = 'https://www.googleapis.com/auth/gmail.compose'
        self.credentials_directory = credentials_directory

    @staticmethod
    def initialize_flags():
        try:
            import argparse
            flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
        except ImportError:
            flags = None
        return flags

    def get_credentials(self):
        credential_dir = self.credentials_directory
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir, 'sendEmail.json')
        store = oauth2client.file.Storage(credential_path)
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

    def get_gmail_api_obj(self):
        credentials = self.get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('gmail', 'v1', http=http)
        return service

    @staticmethod
    def create_message(sender, to, subject, message_text):
        message = MIMEText(message_text, 'html')
        message['to'] = to
        message['from'] = sender
        message['subject'] = subject
        return {'raw': base64.urlsafe_b64encode(message.as_string())}

    @staticmethod
    def send_message(service, user_id, message):
        try:
            message = (service.users().messages().send(userId=user_id, body=message)
                       .execute())
            print 'Message Id: %s' % message['id']
            return message
        except errors.HttpError, error:
            print 'An error occurred: %s' % error

    def send_email(self, from_email, to_emails, subject, message_body):
        service = self.get_gmail_api_obj()
        try:
            message_obj = EmailSender.create_message(from_email, to_emails, subject, message_body)
            EmailSender.send_message(service, "me", message_obj)
        except Exception, e:
            print e
            time.sleep(10)
            sys.exit(1)
'''
if __name__ == "__main__":

    try:
        email = EmailSender('jiraMetrics', "../client_secret.json", "../.credentials")
        service = email.get_gmail_api_obj()
        email.send_message(service, "me", email.create_message("jirametrics.ert@gmail.com", "Siva_Ninala@epam.com, ninalasiva@gmail.com", "Test gmail automation", "Hello <br /> &emsp; world"))

    except Exception, e:
        print e
        raise
'''