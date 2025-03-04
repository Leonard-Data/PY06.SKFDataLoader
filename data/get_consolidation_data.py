from APIs.consolidations import ConsolidationAPI
from models.consolidation import ConsolidationModel
import os
import msal
from msal_requests_auth.auth import ClientCredentialAuth
from dataverse_api import DataverseClient
from requests import Session
import pandas as pd

def get_consolidation_data():
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
    Consolidation = ConsolidationAPI(dataverse_client)

    return Consolidation