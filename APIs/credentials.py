from auth.core import *
from requests import Response
from models.credential import CredentialModel
from dataverse_api import DataverseClient, DataverseError
from typing import Union
from auth.dataverseAPI import get_current_token_dataverse
import argparse

class CredentialAPI:
    def __init__(self, client: DataverseClient=None):
        if client:
            self.client = client
        else:
            self.client: DataverseClient = dataverse_client
        self.entity = self.client.entity(logical_name='cr58d_skfcredential')

    
    def get_credential(self, id: str) -> list[CredentialModel]:
        responses: list[Response] = self.entity.read(select=CredentialModel.get_field_names(by_alias=True),
                                                     filter = f"cr58d_skfcredentialid eq {id}")
        if responses:
            response = responses[0]
            return CredentialModel(**response)
        else:
            return None
                  
    def create_log(self, data: Union[CredentialModel, dict]) -> bool:
        try:
            if not isinstance(data, dict):
                _data = data.dict(by_alias=True, exclude_none=True)
            else:
                model_data = CredentialModel.construct(**data)
                _data = model_data.dict(by_alias=True, exclude_none=True)

            responses: list = self.entity.create(
                [_data],
                detect_duplicates=True,
                threading=True,
                return_created=False
            )
            response: requests.Response = responses[0]
            if response.status_code in [201, 202, 200, 204]:
                return True
            else:
                return False
        except KeyError as e:
            print(f"KeyError: Missing {e} in data.")
            return False
        except ValueError as e:
            print(f"ValueError: {str(e)}")
            return False
        except Exception as e:
            print(f"An unexpected error occurred: {str(e)}")
            return False
