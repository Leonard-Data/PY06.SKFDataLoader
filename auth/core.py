import asyncio
import datetime
import os
import time
import msal
from msal_requests_auth.auth import ClientCredentialAuth
from dataverse_api import DataverseClient
from requests import Session
import requests
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
TENANT_ID = os.environ["TENANT_ID"]
ORG_ENVIRONMENT = os.environ["ORG_ENVIRONMENT"]

authority: str = f"https://login.microsoftonline.com/{TENANT_ID}"
login_redirect="/"
cache = msal.TokenCache()
session = Session()

client_credential = msal.ConfidentialClientApplication(
    client_id= CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=authority,
    token_cache=cache,
)

auth = ClientCredentialAuth(client=client_credential, 
                                    scopes=[ORG_ENVIRONMENT + "/.default offline_access"])
session.auth = auth
dataverse_client: DataverseClient = DataverseClient(session=session, 
                                                environment_url=ORG_ENVIRONMENT)

class AuthState():
    _access_token: str = ""
    _refresh_token: str = ""
    _expiry_time: float = 0.0
    
    
    async def fetch_token(self):
        url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
        headers = {
            "Content-Type": "application/x-www-form-urlencoded"
        }
        """Fetch a new token initially."""
        response = requests.post(url, data={
            'client_id': CLIENT_ID,
            'client_secret': CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default offline_access",
            'grant_type': 'client_credentials'
        }, headers=headers)
        data: dict = response.json()
        self._access_token = data.get('access_token')
        self._refresh_token = data.get('refresh_token')  # Not always provided
        self._expiry_time = time.time() + data.get('expires_in')
    
    async def check_expiry(self) -> bool:
        return time.time() > self._expiry_time - 60*5 # Renew 5 minutes before expiry
    
    async def monitor_token(self):
        """
        Periodically check and renew the token if it's about to expire.
        """
        while True:
            if self.check_expiry():  # Renew 5 minutes before expiry
                await self.fetch_token()
            await asyncio.sleep(60)  # Check every minute
    
    