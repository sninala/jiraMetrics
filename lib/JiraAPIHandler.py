import requests
import urllib
from requests.auth import HTTPBasicAuth


class JiraAPIHandler(object):
    def __init__(self, config):
        self.config = config
        self.base_url = config.get('API', 'search_api_url')
        self.username = config.get('BUG_TRACKER', 'username')
        self.password = self.config.get('BUG_TRACKER', 'password')

    def get_response_from_jira(self, query):
        query_string = urllib.quote_plus(query)
        response = requests.get(self.base_url + query_string + '&maxResults=1',
                                auth=HTTPBasicAuth(self.username, self.password))
        if response.status_code == 200:
            response_json = response.json()
            # writeResponseToFileSystem(project, status, response_json)
            return response_json
        else:
            raise Exception("Unable get response from Jira")
