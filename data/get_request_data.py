from APIs.requests import RequestAPI
from models.request import RequestModel
import os
import msal
from msal_requests_auth.auth import ClientCredentialAuth
from dataverse_api import DataverseClient
from requests import Session
import pandas as pd

def get_request_data():
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
    request = RequestAPI(dataverse_client)

    requests: list[RequestModel] = request.get_codes()
    # Chuyển danh sách các đối tượng RequestModel thành danh sách các từ điển
    data = [item.dict() for item in requests]

    # Chuyển danh sách từ điển thành DataFrame
    df_request = pd.DataFrame(data)
    
    return df_request

def get_dataverse_client():
    CLIENT_ID = os.environ["CLIENT_ID"]
    CLIENT_SECRET = os.environ["CLIENT_SECRET"]
    TENANT_ID = os.environ["TENANT_ID"]
    ORG_ENVIRONMENT = os.environ["ORG_ENVIRONMENT"]

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    cache = msal.TokenCache()
    session = Session()

    # Khởi tạo ConfidentialClientApplication để lấy access token
    client_credential = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority,
        token_cache=cache,
    )

    # Lấy access token
    result = client_credential.acquire_token_for_client(scopes=[ORG_ENVIRONMENT + "/.default"])
    
    if "access_token" not in result:
        raise Exception("Không thể lấy access token từ Microsoft.")
    
    # Thiết lập session với token
    access_token = result["access_token"]
    session.headers.update({
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    })

    # Khởi tạo DataverseClient với session đã có
    dataverse_client = DataverseClient(session=session, environment_url=ORG_ENVIRONMENT)
    return dataverse_client

def update_request_status_remark(request_id: str, status: int, remark: str):
    dataverse_client = get_dataverse_client()
    request_api = RequestAPI(dataverse_client)
    return request_api.update_request_status_remark(request_id, status, remark)

