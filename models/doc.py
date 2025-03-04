from typing import Optional
from pydantic.v1 import BaseModel, Field

class DocModel(BaseModel):
    
    id: Optional[str] = Field(default=None, alias='cr58d_id')
    file_name: str = Field(alias='cr58d_filename')
    path: Optional[str] = Field(default=None, alias='cr58d_path')
    upload_date: str = Field(alias='cr58d_uploaddate')
    request_ID: str = Field(alias='_cr58d_requestid_value')
    
    @classmethod
    def get_field_names(cls, by_alias=False) -> list[str]:
        if by_alias:
            return [field.alias for field in cls.__fields__.values()]
        else:
            return list(cls.__fields__.keys())
        
    class Config:
        allow_population_by_field = True