from typing import Optional
from pydantic.v1 import BaseModel, Field

class CredentialModel(BaseModel):
    
    name: Optional[str] = Field(default=None, alias='cr58d_name')
    value: Optional[str] = Field(default=None, alias='cr58d_value')
    id: Optional[str] = Field(default=None, alias='cr58d_skfcredentialid')
    
    @classmethod
    def get_field_names(cls, by_alias=False) -> list[str]:
        if by_alias:
            return [field.alias for field in cls.__fields__.values()]
        else:
            return list(cls.__fields__.keys())
        
    class Config:
        allow_population_by_field_name = True