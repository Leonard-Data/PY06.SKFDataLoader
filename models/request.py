from typing import Optional
from pydantic.v1 import BaseModel, Field

class RequestModel(BaseModel):
    
    id: Optional[str] = Field(default=None, alias='cr58d_skfrequestid')
    upload_date: str = Field(alias='cr58d_uploaddate')
    modified_date: str = Field(alias='cr58d_modifieddate')
    BU: str = Field(alias='_cr58d_businessunit_value')
    SKF: str = Field(alias='_cr58d_skfkey_value')
    period: Optional[int] = Field(default=None, alias='cr58d_period')
    year: Optional[int] = Field(default=None, alias='cr58d_year')
    company_code: str = Field(alias='cr58d_companycode')
    status_request: str = Field(alias='cr58d_status')
    approval_request: Optional[str] = Field(default=None, alias='cr58d_approval')
    gen_file_ID: Optional[str] = Field(default=None, alias="_cr58d_genfileid_value")
    document_upload_ID: Optional[str] = Field(default=None, alias='_cr58d_documentid_value')
    remark: Optional[str] = Field(default=None, alias='cr58d_remark')
    cr58d_DocumentID: Optional[str] = None
    cr58d_GenFileID: Optional[str] = None 
    
    @classmethod
    def get_field_names(cls, by_alias=False) -> list[str]:
        if by_alias:
            return [field.alias for field in cls.__fields__.values()]
        else:
            return list(cls.__fields__.keys())
        
    class Config:
        allow_population_by_field = True