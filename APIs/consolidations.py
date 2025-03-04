from auth.core import *
from requests import Response
from models.consolidation import ConsolidationModel
from dataverse_api import DataverseClient, DataverseError
from typing import Union

class ConsolidationAPI:
    def __init__(self, client: DataverseClient=None):
        if client:
            self.client = client
        else:
            self.client: DataverseClient = dataverse_client
        self.entity = self.client.entity(logical_name='cr58d_skfconsolidation')

    
    def get_codes_consolidation(self) -> list[ConsolidationModel]:
        responses: list[Response] = self.entity.read(select=ConsolidationModel.get_field_names(by_alias=True),expand=["cr58d_RequestID"])
        for response in responses:
            # Kiểm tra và thay thế giá trị None nếu cần
            if "cr58d_RequestID" in response and not isinstance(response["cr58d_RequestID"], str):
                response["cr58d_RequestID"] = str(response.get("cr58d_RequestID", ""))
            #print(response)
        if responses:
            return [ConsolidationModel(**data) for data in responses]
        else:
            return []
         
         
    def create_consolidation(self, data: Union[ConsolidationModel, dict]) -> bool:
        try:
            if not isinstance(data, dict):
                _data = data.dict(by_alias=True, exclude_none=True)
            else:
                model_data = ConsolidationModel.construct(**data)
                _data = model_data.dict(by_alias=True, exclude_none=True)

            if '_cr58d_requestid_value' in _data:
                _data['cr58d_RequestID@odata.bind'] = f"/cr58d_skfrequests({_data['_cr58d_requestid_value']})"
                del _data['_cr58d_requestid_value']
            else:
                raise ValueError("Missing '_cr58d_requestid_value' in data.")

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

    def delete_by_document_upload_id(self, document_upload_id: str, mode: str = "individual", threading: bool = False) -> list[Response]:
        """
        Delete rows from the Consolidation table where 'document_upload_ID' matches the given value.

        Parameters
        ----------
        document_upload_id : str
            The value of 'document_upload_ID' to match for deletion.
        mode : str
            'individual' for individual DELETE requests, or 'batch' for batch deletion.
        threading : bool
            Whether to use threading for improved performance during deletion.

        Returns
        -------
        list[Response]
            List of responses from the delete operation.
        """
        try:
            filter_condition = f"_cr58d_requestid_value eq '{document_upload_id}'"
            print(filter_condition)
            return self.entity.delete(filter=filter_condition, mode=mode, threading=threading)
        except DataverseError as e:
            print(f"DataverseError: {str(e)}")
            return []
        except Exception as e:
            print(f"An unexpected error occurred: {str(e)}")
            return []