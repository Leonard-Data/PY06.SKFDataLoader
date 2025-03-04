from auth.core import *
from requests import Response
from models.doc import DocModel
from dataverse_api import DataverseClient

class DocAPI:
    def __init__(self, client: DataverseClient=None):
        if client:
            self.client = client
        else:
            self.client: DataverseClient = dataverse_client
        self.entity = self.client.entity(logical_name='cr58d_skfdocument')

    def get_codes(self) -> list[DocModel]:
        responses: list[Response] = self.entity.read(select=DocModel.get_field_names(by_alias=True))
        if responses:
            return [DocModel(**data) for data in responses]
        else:
            return []
        
    def add_code(self, code: DocModel) -> bool:
        """
        Create a new record in Dataverse.
        """
        data = code.model_dump(by_alias=True, 
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