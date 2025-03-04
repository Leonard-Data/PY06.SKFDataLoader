from typing import Optional
from pydantic.v1 import BaseModel, Field

class LogModel(BaseModel):
    
    id: Optional[str] = Field(default=None, alias='cr58d_skflogid')
    messages: Optional[str] = Field(default=None, alias='cr58d_message')
    title: Optional[str] = Field(default=None, alias='cr58d_title')
    timestamp: Optional[str] = Field(default=None, alias='cr58d_timestamp')
    request_id: Optional[str] = Field(default=None, alias='_cr58d_requestid_value')
    
    @classmethod
    def get_field_names(cls, by_alias=False) -> list[str]:
        if by_alias:
            return [field.alias for field in cls.__fields__.values()]
        else:
            return list(cls.__fields__.keys())
        
    class Config:
        allow_population_by_field_name = True