from typing import Optional
from pydantic.v1 import BaseModel, Field

class ConsolidationModel(BaseModel):
    
    id: Optional[str] = Field(default=None, alias='cr58d_id')
    document_upload_ID: Optional[str] = Field(default=None, alias='_cr58d_requestid_value')
    company_code: str = Field(alias='cr58d_companycode')
    quantity: int = Field(alias='cr58d_quantity')
    remark: Optional[str] = Field(default=None, alias='cr58d_remark')
    return_document_no: Optional[str] = Field(default=None, alias='cr58d_returndocumentno')
    messages: Optional[str] = Field(default=None, alias='cr58d_messages')
    cr58d_RequestID: Optional[str] = None
    period: Optional[str] = Field(default=None, alias='cr58d_period')
    year: Optional[str] = Field(default=None, alias='cr58d_year')
    return_revert_document_no: Optional[str] = Field(default=None, alias='cr58d_returnrevertdocumentno')
    revert_messages: Optional[str] = Field(default=None, alias='cr58d_revertmessages')
    consolidation_id: Optional[str] = Field(default=None, alias='cr58d_skfconsolidationid')
    
    @classmethod
    def get_field_names(cls, by_alias=False) -> list[str]:
        if by_alias:
            return [field.alias for field in cls.__fields__.values()]
        else:
            return list(cls.__fields__.keys())
        
    class Config:
        allow_population_by_field = True