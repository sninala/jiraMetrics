import requests
import urllib
from requests.auth import HTTPBasicAuth


class JiraAPIHandler(object):
    def __init__(self, config):
        self.config = config
        self.base_url = config.get('API', 'search_api_url')
        self.username = config.get('BUG_TRACKER', 'username')
        self.password = self.config.get('BUG_TRACKER', 'password')

    def get_response_from_jira(self, query, start_at=None, fields=None):
        query_string = urllib.quote_plus(query)
        if start_at and fields:
            request_url = self.base_url + query_string + '&maxResults=-1&startAt='+ start_at +'&fields='+fields
        else:
            request_url = self.base_url + query_string + '&maxResults=1'
        response = requests.get(request_url, auth=HTTPBasicAuth(self.username, self.password))
        if response.status_code == 200:
            response_json = response.json()
            return response_json
        else:
            raise Exception("Unable get response from Jira - {}".format(response.reason))
