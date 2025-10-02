from auth.core import *
from requests import Response
from models.request import RequestModel
from dataverse_api import DataverseClient
from auth.dataverseAPI import get_current_token_dataverse
from typing import Optional
from utils.retry import retry_with_backoff
import logging
import requests
from requests.exceptions import RequestException, Timeout, ConnectionError

logger = logging.getLogger(__name__)

class RequestAPI:
    def __init__(self, client: Optional[DataverseClient]=None):
        self.timeout = 180.0  # 3 minutes
        if client:
            self.client = client
        else:
            self.client: DataverseClient = dataverse_client
        self.entity = self.client.entity(logical_name='cr58d_skfrequest')

    @retry_with_backoff(
        max_retries=3,
        base_delay=30.0,
        timeout=180.0,
        exceptions=(RequestException, Timeout, ConnectionError, Exception)
    )
    def get_codes(self) -> list[RequestModel]:
        """Get codes with automatic retry"""
        logger.info("Fetching data codes...")
        expand_fields = "cr58d_DocumentID,cr58d_GenFileID"
        responses: list[dict] = self.entity.read(select=RequestModel.get_field_names(by_alias=True),expand=expand_fields)
        
        for response in responses:
            # Kiểm tra và thay thế giá trị None nếu cần
            if "cr58d_DocumentID" in response and not isinstance(response["cr58d_DocumentID"], str):
                response["cr58d_DocumentID"] = str(response.get("cr58d_DocumentID", ""))
            if "cr58d_GenFileID" in response and not isinstance(response["cr58d_GenFileID"], str):
                response["cr58d_GenFileID"] = str(response.get("cr58d_GenFileID", ""))

        if responses:
            return [RequestModel(**data) for data in responses]
        else:
            return []
        
    def add_code(self, code: RequestModel) -> bool:
        """
        Create a new record in Dataverse.
        """
        data = code.dict(by_alias=True, 
                         exclude_none=True)

        responses: list = self.entity.create([data],
                                      detect_duplicates=True,
                                      threading=True,
                                      return_created=False)
        response: requests.Response = responses[0]
        if response.status_code in [201,202,200,204]:
            return True
        else: 
            return False
        
    def update_request_status_remark(self, request_id: str, status: int, remark: str) -> bool:
        """
        Cập nhật trạng thái của một yêu cầu (request) trong Dataverse.
        """
        try:
            # Dữ liệu cần cập nhật
            data = {
                'cr58d_status': status,
                'cr58d_remark': remark,
            }
            base_url = "https://faportals.crm5.dynamics.com/api/data/v9.2"
            update_url = f"{base_url}/cr58d_skfrequests({request_id})"
            
            # Lấy access token từ msal (sử dụng client_credential)
            token = get_current_token_dataverse()
            
            # Đảm bảo token đã được lấy thành công
            if not token:
                print("Failed to get access token.")
                return False
            
            # Cập nhật headers với token
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }

            # Gửi yêu cầu PATCH để cập nhật trạng thái
            response: Response = requests.patch(update_url, json=data, headers=headers)

            # Kiểm tra mã trạng thái phản hồi
            if response.status_code in [200, 204]:
                print(f"Successfully updated status for request {request_id}")
                return True
            else:
                print(f"Failed to update status for request {request_id}. Response: {response.status_code}, Content: {response.text}")
                return False
        except Exception as e:
            print(f"An error occurred while updating the status: {e}")
            return False
