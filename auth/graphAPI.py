import threading
import time
import requests
import os
from dotenv import load_dotenv
import atexit
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

load_dotenv()

client_id = os.environ["CLIENT_ID"]
client_secret = os.environ["CLIENT_SECRET"]
tenant_id = os.environ["TENANT_ID"]
scope = "https://graph.microsoft.com/.default"

# Global variables for token and expiry time
access_token = None
token_expiry_time = 0
lock = threading.Lock()
stop_event = threading.Event()
should_stop = False  # Signal to stop the thread after main.py completes

def get_access_token(tenant_id, client_id, client_secret, scope):
    """
    Authenticate using OAuth2 to obtain an access token from Azure AD.
    """
    global access_token, token_expiry_time
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope
    }

    response = requests.post(url, headers=headers, data=data)
    
    if response.status_code == 200:
        with lock:
            access_token = response.json().get('access_token')
            expires_in = response.json().get('expires_in')
            token_expiry_time = time.time() + expires_in - 60  # Refresh 1 minute before expiration
            print(f"New graph token created, Expires in: {expires_in} seconds")
        return access_token, token_expiry_time
    else:
        print(f"Failed to obtain access token. Status code: {response.status_code}, Response: {response.text}")
        return None, None

def get_sharepoint_context(site_url):
    """
    Create and return a SharePoint ClientContext object.
    """
    global client_id, client_secret
    client_credentials = ClientCredential(client_id, client_secret)
    ctx = ClientContext(site_url).with_credentials(client_credentials)
    return ctx

def monitor_token():
    global access_token, token_expiry_time
    while not stop_event.is_set():  # Gently check for stop signal
        if time.time() >= token_expiry_time:
            print("Graph token expired, refreshing...")
            access_token, token_expiry_time = get_access_token(tenant_id, client_id, client_secret, scope)
            print("New graph token obtained")
        else:
            pass
            #print("Graph Token is still valid, no refresh needed.")
        time.sleep(10)

def stop_monitoring():
    global should_stop
    print("Stopping graph token monitoring...")
    stop_event.set()  # Update stop signal when main.py completes


# Function to get the current token
def get_current_token():
    with lock:
        return access_token
