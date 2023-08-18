from dataclasses import dataclass, field
from abc import ABC, abstractmethod

from typing import List
import pytz
from datetime import datetime
import msal
import requests

class SharePointError(Exception):
    '''Exception raised if input json value has empty rows'''

class InvalidResponseCodeError(SharePointError):
    def __init__(self, response_code: str, failure: str):
        super().__init__(f'Invalid response code: {response_code}!')

class Validator(ABC):
    '''
    Validator abstract class
    '''
    def __get__(self, obj, objtype=None):
        return self.value
    def __set__(self, obj, value):
        self.validate(value)
        self.value = value
    @abstractmethod
    def validate(self, value):
        pass

class Response(Validator):
    '''
    Http response code validator
    '''
    valid_response = (requests.codes.ok, 201)
    def validate(self, response):
        if response.status_code not in Response.valid_response:
            raise InvalidResponseCodeError(
                response_code = response.status_code,
                failure=''
            )

@dataclass
class MicrosoftGraph:
    """
    class used to authenticate MicrosoftGraph Connection and query items

    ...
    Attributes
    ----------

    client_id: str
        the client ID is the unique application (client) ID assigned to your app by Azure AD when the app was registered.

    tenant_id: str
        the tenant ID, a unique identifier representing the organization (or tenant)

    client_credential: str
        this is your Graph API Client ID.

    username: str
        username with appropriate permissions

    password: str
        associated with the username

    response: str = field(init=False, default=Response())
        the response from the API call

    scope: List[str]
        the scope of the API call

    Methods
    -------
    _get_token()
        generates token

    get_sharepoint_list_item_data(self, tenant_name: str, team_id: str, list_id: str, column: str, value: str)
        retries data from a sharepoint list with a team on sharepoint

    """
    client_id: str
    tenant_id: str
    client_credential: str
    username: str #might be your email
    password: str
    response: str = field(init=False, default=Response())
    scope: List[str] = field(default_factory=lambda: ['https://graph.microsoft.com/.default'])
    app: msal.ClientApplication = None

    def __init__(self, client_id: str, tenant_id: str, client_credential: str, username: str, password: str, scope: List[str]):
        
        """
        Create an instance of the msal application

        Parameters
        ----------
        client_id: str
            the client ID is the unique application (client) ID assigned to your app by Azure AD when the app was registered.

        tenant_id: str
            the tenant ID, a unique identifier representing the organization (or tenant)

        client_credential: str
            this is your Graph API Client ID.

        username: str
            Username with appropriate permissions

        password: str
            associated with the username

        scope: List[str]
            the scope of the API call
        """

        self.client_id = client_id
        self.tenant_id = tenant_id
        self.client_credential = client_credential
        self.username = username
        self.password = password
        self.scope = scope

        if self.app == None:
    
            self.app = msal.ClientApplication(
            
                client_id = self.client_id, 
                authority = f"https://login.microsoftonline.com/{self.tenant_id}/",
                client_credential = self.client_credential
            )
    
    def _get_token(self) -> str:
        '''
        Grab Graph API token leveraging MSAL library
        '''
        result = None
        accounts = self.app.get_accounts(username=self.username)
        # print(accounts)
        if accounts:
            #print("Account(s) exists in cache, probably with token too. Let's try.")
            result = self.app.acquire_token_silent(self.scope, account=accounts[0])


        if not result:
            #print("No suitable token exists in cache. Let's get a new one from AAD.")
           
            result = self.app.acquire_token_by_username_password(
                self.username, 
                self.password, 
                scopes=self.scope
            )
        if 'access_token' in result:
            return result['access_token']

    def get_sharepoint_list_item_data(self, tenant_name: str, team_id: str, list_id: str, column: str, value: str) -> requests.models.Response:
        
        """returns data from a sharepoint list with a team on sharepoint

        Parameters
        ----------
        
        tenant_name: str
            The tenant name, a unique identifier representing the organization (or tenant) 

        team_id: str
            the sharepoint team which contains the list

        list_id: str
            the unique id of the list

        column: str
            the column name off of which the row is extracted

        value: str
            the value of the column

        """
        try:
            url = f'https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/teams/{team_id}:/lists/{list_id}/items?$expand=fields&$filter=fields/{column} eq \'{value}\''

            headers={
                'Authorization': f'Bearer {self._get_token()}'
            }

            self.response = requests.get(url = url, headers = headers)

            if self.response.status_code == 200:
                return self.response
        except Exception as e:
            timezone = pytz.timezone('Australia/Sydney')
            current_time = datetime.now(timezone)
            print(f"SharePoint connection error at: {current_time} with error {e}")

    def get_sharepoint_list(self, tenant_name: str, team_id: str, list_id: str) -> requests.models.Response:
        
        """returns data from a sharepoint list with a team on sharepoint - this is a entire list

        Parameters
        ----------
        
        tenant_name: str
            The tenant name, a unique identifier representing the organization (or tenant) 

        team_id: str
            the sharepoint team which contains the list

        list_id: str
            the unique id of the list

        """
        try:
            url = f'https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/teams/{team_id}:/lists/{list_id}/items'
            print(url)
            headers={
                'Authorization': f'Bearer {self._get_token()}'
            }

            self.response = requests.get(url = url, headers = headers, params={'expand': 'False', 'top': '4999'})

            if self.response.status_code == 200:
                return self.response
        except Exception as e:
            timezone = pytz.timezone('Australia/Sydney')
            current_time = datetime.now(timezone)
            print(f"SharePoint connection error at: {current_time} with error {e}")