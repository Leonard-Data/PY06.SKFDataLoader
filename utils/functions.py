from configparser import ConfigParser
import subprocess
import requests
from models.consolidation import ConsolidationModel
from APIs.consolidations import ConsolidationAPI
from . import setup_logger
import pandas as pd
from auth.graphAPI import get_current_token
import argparse
import logging
import os
import threading
import time
import json
import pandas as pd
import xlsxwriter
import shutil
import sys
import pytz
import math
import csv
from datetime import datetime, timedelta, timezone
from logging import config
from openpyxl import load_workbook
from xlsx2csv import Xlsx2csv
from utils.functions import *
from utils.log import setup_logger
from requests.exceptions import RequestException
from urllib.parse import quote
import concurrent.futures
from typing import Tuple
from requests import post

logger = setup_logger(os.path.basename(__file__))

def kill_process(process_names: list[str])-> None:     
    # Replace 'sapgui.exe' with the actual process name of your SAP application
    for name in process_names:
        try:
            subprocess.run(f"taskkill /f /im {name}", shell=False, check=True)
            logger.info(f"Successfully killed {name}")
        except subprocess.CalledProcessError as e:
            logger.error(f"Error: {e}")

def get_access_token(tenant_id, client_id, client_secret, scope) -> str | None:
        """
        Authenticate using OAuth2 to get an access token from Azure AD.

        Args:
        - tenant_id (str): The Azure AD tenant ID.
        - client_id (str): The client ID of the registered application.
        - client_secret (str): The client secret of the registered application.
        - scope (str): The scope for which the token is requested, usually the resource URL with '/.default'.

        Returns:
        - str: The access token.
        """
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        headers = {
            "Content-Type": "application/x-www-form-urlencoded"
        }
        data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": scope
        }
        
        response = requests.post(url, headers=headers, data=data)
        
        if response.status_code == 200:
            token = response.json().get('access_token')
            return token
        else:
            raise Exception(f"Failed to obtain access token. Status Code: {response.status_code}, Response: {response.text}")

def retrieve_credential(vault_path: str, token: str, name: str):
    header = {
            'Authorization': f'{token}',
        }
    payload = {}
    try:
            url = f"{vault_path}/{name}:/content"
            response = requests.request("GET", url, headers=header, data=payload)
            return response.json()
        
    except requests.exceptions as err:
        logger.error(f"Unexpected Error: {err}")
        return None
        # logger.error(f"HTTP error occurred: {err}")

def get_credentials() -> requests.Response | None:
    base_path = os.path.dirname(__file__).replace("utils","")
    config = ConfigParser()
    config.read([os.path.join(base_path,"config.cfg"), 
                os.path.join(base_path,"config.dev.cfg")])
    token = get_access_token(tenant_id=config['Graph']['tenantId'],
            client_id=config['Graph']['clientId'],
            client_secret=config['Graph']['clientSecret'],
            scope=config['Graph']['scopes'])
    vault_path = f"https://graph.microsoft.com/v1.0/sites/{config['Sharepoint']['siteID']}/drives/{config['Sharepoint']['driveID']}/{config['Sharepoint']['VaultPath']}"
    return retrieve_credential(vault_path,token=token,name=config['Vault']['Name'])

def clear_folder(destination: str):
    if os.path.isdir(destination):
          for root, dirs, files in os.walk(destination):
            for f in files:
                os.unlink(os.path.join(root, f))
            for d in dirs:
                shutil.rmtree(os.path.join(root, d))
    else:
         os.makedirs(destination)

def df_to_consolidation_models(df: pd.DataFrame) -> list[ConsolidationModel]:
    models = []
    for _, row in df.iterrows():
        # Chuyển từng hàng thành từ điển
        data = row.to_dict()

        # Khởi tạo ConsolidationModel từ từ điển
        model = ConsolidationModel(**data)
        models.append(model)
    return models

def consolidation_models_to_df ():
    
    consolidation_api = ConsolidationAPI() 
    Consolidations: list[ConsolidationModel] = consolidation_api.get_codes_consolidation()
    # Chuyển danh sách các đối tượng RequestModel thành danh sách các từ điển
    data = [item.dict() for item in Consolidations]

    # Chuyển danh sách từ điển thành DataFrame
    df_consolidation = pd.DataFrame(data)
    
    return df_consolidation

def upload_dataframe_to_consolidation(df: pd.DataFrame, consolidation_api: ConsolidationAPI)-> list[dict]:
    models = df_to_consolidation_models(df)
    upload_results = []
    
    for model in models:
        success = consolidation_api.create_consolidation(model)
        if success:
            status = "Successful"
            logger.info(f"Uploaded: {model}")
        else:
            status = "Failed"
            logger.error(f"Failed to upload: {model}")
        
        upload_results.append({
            'model': model,
            'status_upload': status
        })

    return upload_results 
   
def process_consolidation_data(input_file_path: str) -> pd.DataFrame:

    # Step 1: Read data from the Excel file
    df = pd.read_excel(input_file_path, sheet_name='Sheet1')
    
    # Step 2: Slice the first 4 characters from the 'Document Header Text' column
    df['Document Header Text'] = df['Document Header Text'].str.slice(0, 4)

    # Step 3: Remove rows containing NaN values
    df_result = df.dropna(subset=['Document Header Text'])
   
    df_grouped = df
    grouped_columns_needed = ['Document Header Text', 'Total Quantity']
    df_grouped = df_grouped[grouped_columns_needed]
    df_grouped['Document Header Text'] = df_grouped['Document Header Text'].ffill()
    
    # Step 4: Convert the 'Total Quantity' column to a numeric type (float or int)
    df['Total Quantity'] = pd.to_numeric(df['Total Quantity'], errors='coerce').round(2)
    df_grouped['Total Quantity'] = pd.to_numeric(df_grouped['Total Quantity'], errors='coerce').round(2)
    
    # Step 5: Group by 'Document Header Text' and sum the 'Total Quantity' column from the original dataframe
    df_grouped = df_grouped.groupby('Document Header Text', as_index=False)['Total Quantity'].sum()
    df_grouped['Total Quantity'] = df_grouped['Total Quantity'].round(2)
    # Step 6: Merge the result back into df_result
    df_result = df_result.merge(df_grouped, on='Document Header Text', how='left')
    
    # Step 7: Filter the required columns
    columns_needed = ['Document Header Text', 'Message Text', 'Total Quantity_y']
    df_result = df_result[columns_needed]

    # Step 8: Create a new column 'cr58d_returndocumentno' by extracting the last 10 characters from 'Message Text'
    df_result['cr58d_returndocumentno'] = df_result['Message Text'].apply(
        lambda x: x[-10:] if 'Document is posted under number' in str(x) else None
    )

    # Step 9: Map 'Document Header Text' values to numbers
    header_text_map = {
        'VN81': 0,
        'VN82': 1,
        'VN83': 2,
        'VN87': 3,
        'VN88': 4,
        'VN89': 5,
        'VN91': 6,
        'VN92': 7,
        'VN9B': 8,
        'VNZ1': 9
    }

    def map_header_text(value):
        if value in header_text_map:
            return header_text_map[value]
        else:
            error_message = f"Error: '{value}' is not in the company code mapping table. Stopping the process."
            logger.error(error_message)
            raise Exception(error_message)  # Stop the process and raise an error

    # Step 10: Apply the mapping function to 'Document Header Text'
    df_result['Document Header Text'] = df_result['Document Header Text'].apply(map_header_text)

    # Step 11: Rename columns
    df_result = df_result.rename(columns={
        'Document Header Text': 'cr58d_companycode',
        'Message Text': 'cr58d_messages',
        'Total Quantity_y': 'cr58d_quantity'
    })

    # Step 12: Add new columns with default values as None or any other required values
    df_result['_cr58d_requestid_value'] = None
    df_result['cr58d_id'] = None
    df_result['cr58d_remark'] = None

    total_rows = len(df_result)
    return_doc_count = df_result['cr58d_returndocumentno'].notna().sum()
    
    if return_doc_count == 0:
        status = "Failed"
    elif 0 < return_doc_count < total_rows:
        status = "Failed. Revert as only partial import succeeded."
    elif return_doc_count == total_rows:
        status = "Completed"
    else:
        status = "Unknown"

    # Return DataFrame and status
    return df_result, status

def get_credentials():
    try:
        
        with open('credentials.json', 'r') as f:
            credentials = json.load(f)
        if not credentials:
            raise ValueError("The credentials file is empty.")
        return credentials
    except json.JSONDecodeError as e:
        logger.error(f"Error decoding JSON: {e}")
        raise
    except FileNotFoundError:
        logger.error("The credentials.json file was not found.")
        raise
    except ValueError as e:
        logger.error(f"Error: {e}")
        raise
    
def create_folder():
    # Create a folder to store downloaded files
    folder_path = "C:\\Users\\Public\\Downloads\\ACP.RefreshData"

    # Check the folder
    if not os.path.exists(folder_path):
        # Create a new folder if it does not exist
        os.makedirs(folder_path)
        logger.info(f"Folder '{folder_path}' has been created.")
    else:
        # If the folder exists, delete the files inside
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # Delete file or link
                    logger.info(f"Deleted file: {file_path}")
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # Delete subdirectory
                    logger.info(f"Deleted folder: {file_path}")
            except Exception as e:
                logger.error(f"Failed to delete {file_path}. Reason: {e}")

        logger.info(f"All files in '{folder_path}' have been cleared.")
    return folder_path

def process_log_data(df_result: pd.DataFrame, df_consolidation: pd.DataFrame) -> str:
    # Step 1: Create "Mapping_result" and "Mapping_consolidation"
    df_result['Mapping_result'] = df_result['cr58d_companycode'].astype(str) + '_' + df_result['_cr58d_requestid_value'].astype(str)
    df_consolidation['Mapping_consolidation'] = df_consolidation['company_code'].astype(str) + '_' + df_consolidation['document_upload_ID'].astype(str)
    
    # Step 2: Create "cr58d_skfconsolidationid" in df_result and use vlookup-like method to assign values
    df_result = df_result.merge(df_consolidation[['Mapping_consolidation', 'consolidation_id']], 
                                left_on='Mapping_result', right_on='Mapping_consolidation', 
                                how='left')
    df_result['cr58d_skfconsolidationid'] = df_result['consolidation_id']
                
    header_text_map = {
        0 : 'VN81',
        1 : 'VN82' ,
        2 : 'VN83',
        3 : 'VN87',
        4 : 'VN88',
        5 : 'VN89',
        6 : 'VN91',
        7 : 'VN92',
        8 : 'VN9B',
        9 : 'VNZ1'
    }

    def map_header_text(value):
        if value in header_text_map:
            return header_text_map[value]
        else:
            error_message = f"Error: '{value}' is not in the company code mapping table. Stopping the process."
            logger.error(error_message)
            raise Exception(error_message)  # Stop the process and raise an error

    df_result['cr58d_companycode'] = df_result['cr58d_companycode'].apply(map_header_text)
    
    # Step 3: Rename columns
    df_result.rename(columns={'cr58d_messages': 'cr58d_revertmessages', 
                              'cr58d_returndocumentno': 'cr58d_returnrevertdocumentno'}, 
                     inplace=True)
    
    # Step 4: Keep only the required columns
    df_result = df_result[['cr58d_companycode', 'cr58d_revertmessages']]
    
    # Remove duplicate rows
    df_result.drop_duplicates(inplace=True)

    # Create the text with new line between each pair of values
    result_text = "\n".join(f"{row['cr58d_companycode']} : {row['cr58d_revertmessages']}" 
                           for _, row in df_result.iterrows())
    
    return result_text

def convert_excel_to_csv_with_formatting(excel_path, csv_path):
    temp_csv_path = csv_path.replace(".CSV", "_temp.CSV")
    try:
        Xlsx2csv(excel_path, outputencoding="utf-8-sig").convert(temp_csv_path)
        
        with open(temp_csv_path, "r", encoding="utf-8-sig") as infile, open(csv_path, "w", newline="", encoding="utf-8-sig") as outfile:
            reader = csv.reader(infile)
            writer = csv.writer(outfile, quoting=csv.QUOTE_NONNUMERIC)

            header = next(reader)
            writer.writerow(header)

            target_cols = ["Period", "Fiscal Year", "Client Code", "Cost Center"]
            col_indices = [i for i, col in enumerate(header) if col in target_cols]
            
            for row in reader:
                for i in col_indices:
                    if row[i].strip():
                        row[i] = f"'{row[i]}"
                writer.writerow(row)

        os.remove(temp_csv_path)
        return True
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        return False
    
def Consolidation_report_masterdata(full_report_file_path, full_masterdata_file_path):
    try:
        df_report = pd.read_excel(full_report_file_path, dtype=str)
        new_columns = ["Client", "Client Code", "Business Line", "BU", "Business Model", "Local 1", "General Name"]
        for col in new_columns:
            df_report[col] = None

        df_masterdata = pd.read_excel(full_masterdata_file_path, sheet_name="Copy", dtype=str, header=0)
        df_masterdata.columns = df_masterdata.columns.str.strip()
        df_report.columns = df_report.columns.str.strip()

        cost_center_column_name = next((col for col in df_masterdata.columns if "cost center" in col.lower()), None)
        if not cost_center_column_name:
            return "Error: The 'cost center' column was not found in the MasterData file!"

        df_masterdata = df_masterdata[df_masterdata[cost_center_column_name] != "#"]
        df_masterdata[cost_center_column_name] = df_masterdata[cost_center_column_name].astype(str).str.strip()
        df_report["Cost Center"] = df_report["Cost Center"].astype(str).str.strip()

        column_mapping = {
            "Client": "Client",
            "Client Code": "Client Code",
            "Business Line": "Business Line 1",
            "Business Model": "Business Model 1",
            "Local 1": "Local 1",
            "General Name": "General Name"
        }
        valid_columns = [col for col in column_mapping.values() if col in df_masterdata.columns]

        if not valid_columns:
            return "Error: No valid columns found in the MasterData file!"

        lookup_dict = df_masterdata.set_index(cost_center_column_name)[valid_columns].to_dict(orient="index")

        for col in column_mapping.keys():
            if column_mapping[col] in valid_columns:
                df_report[col] = df_report["Cost Center"].map(lambda x: lookup_dict.get(x, {}).get(column_mapping[col], "N/A"))

        def get_bu(cost_center):
            if len(cost_center) >= 3:
                return {
                    "1": "CG",
                    "3": "HEC",
                    "4": "PM",
                    "5": "TEC",
                    "9": "OTH"
                }.get(cost_center[2], "NONE")
            return "NONE"

        df_report["BU"] = df_report["Cost Center"].apply(get_bu)

        # Tạo tên file đầu ra .parquet
        report_parquet_filename = os.path.splitext(os.path.basename(full_report_file_path))[0] + ".parquet"
        output_parquet_path = os.path.join(os.path.dirname(full_report_file_path), report_parquet_filename)

        # Ghi file .parquet
        df_report.to_parquet(output_parquet_path, index=False)
        logger.info(f"Successfully saved report to {output_parquet_path}")

        return output_parquet_path

    except Exception as e:
        return f"Error: {str(e)}"
    
def ProcessMasterData(full_masterdata_file_path):
    try:
        temp_file = full_masterdata_file_path.replace(".xlsx", "_temp.xlsx")
        shutil.copy(full_masterdata_file_path, temp_file)

        try:
            wb = load_workbook(full_masterdata_file_path, read_only=True)
            if "PCC" not in wb.sheetnames:
                return "Error: The 'PCC' sheet is missing from the MasterData file!"
            ws_pcc = wb["PCC"]
        except Exception as e:
            return f"Error while loading workbook: {str(e)}"

        try:
            data = list(ws_pcc.iter_rows(values_only=True))
            df = pd.DataFrame(data)
            header = df.iloc[0]  
            df = df[1:].reset_index(drop=True)  

            columns_to_keep = [14, 15, 18, 23, 25, 29, 30]
            df = df.iloc[:, columns_to_keep]
            df.columns = header[columns_to_keep]
            df.columns = df.columns.to_list()[:-1] + ["Client Code"]
            df.columns = [col if col != df.columns[1] else "General Name" for col in df.columns]
        except Exception as e:
            return f"Error while processing data: {str(e)}"

        output_file = full_masterdata_file_path.replace(".xlsx", "_optimized.xlsx")

        try:
            with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name="Copy", index=False, header=True)
        except Exception as e:
            return f"Error while writing to Excel: {str(e)}"

        try:
            shutil.move(output_file, full_masterdata_file_path)
        except Exception as e:
            return f"Error while replacing the original file: {str(e)}"

        return full_masterdata_file_path

    except Exception as e:
        return f"Error: {str(e)}"

def get_report_SKF(sap, month, year, folder_path, file_path):
    max_retries = 2
    for attempt in range(max_retries):
        try:
            # Run transaction YFI_OC_GEN_R027
            sap.run_transaction("YFI_OC_GEN_R027")
            sap.input_text("wnd[0]/usr/ctxtS_KOKRS-LOW", "VN99")
            sap.input_text("wnd[0]/usr/ctxtS_PERIO-LOW", month)
            sap.input_text("wnd[0]/usr/txtS_GJAHR-LOW", year)
            # sap.input_text("wnd[0]/usr/ctxtS_STAGR-LOW", "CRL100")
            sap.click_button("wnd[0]/tbar[1]/btn[8]")  # Click to continue

            # Select radio button and enter path/file name
            sap.select_radio_button("wnd[0]/mbar/menu[0]/menu[3]/menu[1]")
            sap.input_text("wnd[1]/usr/ctxtDY_PATH", folder_path)
            sap.input_text("wnd[1]/usr/ctxtDY_FILENAME", file_path)
            sap.click_button("wnd[1]/tbar[0]/btn[0]")  # Click Generate button
            kill_process(["Excel.exe"])

            if "already exists" in sap.read_log():
                sap.click_button("wnd[1]/tbar[0]/btn[11]")  # Click Replace if exists
                kill_process(["Excel.exe"])

            # Check if file has been downloaded
            if os.path.exists(os.path.join(folder_path, file_path)):
                logger.info(f"File {file_path} downloaded successfully.")
                return True
            else:
                if attempt < max_retries - 1:
                    logger.warning(f"File {file_path} not found. Retrying download...")
                    time.sleep(10)  # Wait before retrying
                else:
                    logger.error(f"Failed to download {file_path} after {max_retries} attempts.")
                    return False

        except Exception as e:
            logger.error(f"Error during report generation (attempt {attempt + 1}): {str(e)}")
            if attempt == max_retries - 1:
                logger.error("Max retries reached. Report generation failed.")
                return False
            time.sleep(5)  # Wait before retrying

    return False

def check_sharepoint_access(access_token, drive_id):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        logger.info("SharePoint access verified.")
        return True
    else:
        logger.error(f"Failed to access SharePoint. Status code: {response.status_code}")
        logger.error(f"Response: {response.text}")
        return False
 
def check_sharepoint_folder(access_token, drive_id, folder_path):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    encoded_folder_path = quote(folder_path)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{encoded_folder_path}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        folder_info = response.json()
        logger.info(f"SharePoint folder info: {json.dumps(folder_info, indent=2)}")
        return True
    else:
        logger.error(f"Failed to get SharePoint folder info. Status code: {response.status_code}")
        logger.error(f"Response: {response.text}")
        return False
       
def upload_file_to_sharepoint(file_path, sharepoint_folder_path):
    max_retries = 3
    chunk_size = 10 * 1024 * 1024  # 10 MB chunks
    
    for attempt in range(max_retries):
        try:
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path)
            drive_id = "b!j5B2Hm8LOES-gp19kZLFNy_qoT96WGBHixaNdzAdCqJpL-7oslv5RI-IsvDE2lkX"
            
            # Kiểm tra xem file có tồn tại không
            check_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}"
            headers = {
                'Authorization': f'Bearer {get_current_token()}',
                'Content-Type': 'application/json'
            }
            response = requests.get(check_url, headers=headers)
            if response.status_code == 200:
                logger.info(f"File '{file_name}' already exists on SharePoint.")
                # File tồn tại, xóa nó
                delete_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}"
                delete_response = requests.delete(delete_url, headers=headers)
                logger.info(f"Delete response status code: {delete_response.status_code}")
                logger.info(f"Delete response content: {delete_response.text}")
                if delete_response.status_code in [204, 404]:
                    logger.info(f"Existing file '{file_name}' deleted successfully.")
                else:
                    logger.error(f"Failed to delete existing file '{file_name}'. Status code: {delete_response.status_code}")
                    logger.error(f"Response content: {delete_response.text}")
                    # Không raise Exception, tiếp tục với upload
            elif response.status_code == 404:
                logger.info(f"File '{file_name}' does not exist on SharePoint.")
            else:
                logger.error(f"Unexpected response when checking file existence. Status code: {response.status_code}")
                logger.error(f"Response content: {response.text}") 
                
            # Tạo phiên upload với tham số conflictBehavior trong URL
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}:/createUploadSession?@microsoft.graph.conflictBehavior=replace"
            body = {
                "item": {
                    "@microsoft.graph.conflictBehavior": "replace"
                }
            }
            response = requests.post(url, headers=headers, json=body)
            response.raise_for_status()
            upload_url = response.json().get('uploadUrl')
            if not upload_url:
                raise Exception(f"Failed to create upload session: {response.text}")
            
            # Upload file theo chunks
            with open(file_path, 'rb') as file:
                for i in range(0, file_size, chunk_size):
                    chunk = file.read(chunk_size)
                    start = i
                    end = min(i + chunk_size - 1, file_size - 1)
                    
                    headers = {
                        'Content-Length': str(len(chunk)),
                        'Content-Range': f'bytes {start}-{end}/{file_size}'
                    }
                    response = requests.put(upload_url, headers=headers, data=chunk)
                    
                    if response.status_code not in [200, 201, 202, 204]:
                        raise Exception(f"Upload failed: {response.text}")
            
            logger.info(f"File uploaded successfully: {file_name}")
            return True
        
        except Exception as e:
            logger.error(f"Error during upload (attempt {attempt + 1}): {str(e)}")
            if "file is currently checked out" in str(e).lower() or "file is locked" in str(e).lower():
                logger.warning("File is currently in use. Waiting before retry...")
                time.sleep(30)  # Wait for 30 seconds before retrying
            elif attempt == max_retries - 1:
                logger.error("Max retries reached. Upload failed.")
                return False
            else:
                time.sleep(5)  # Wait before retrying for other errors

    return False

def download_file_from_sharepoint(sharepoint_file_masterdata_path, folder_path):
    max_retries = 3
    drive_id = "b!j5B2Hm8LOES-gp19kZLFNy_qoT96WGBHixaNdzAdCqJpL-7oslv5RI-IsvDE2lkX"

    file_name = os.path.basename(sharepoint_file_masterdata_path)
    local_file_path = os.path.join(folder_path, file_name)

    for attempt in range(max_retries):
        try:
            # Check if file exists on SharePoint
            check_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_file_masterdata_path}"
            headers = {
                'Authorization': f'Bearer {get_current_token()}',
                'Content-Type': 'application/json'
            }
            response = requests.get(check_url, headers=headers)

            if response.status_code == 200:
                logger.info(f"File '{file_name}' exists on SharePoint, proceeding with download.")
            elif response.status_code == 404:
                logger.error(f"File '{file_name}' does not exist on SharePoint.")
                return False
            else:
                logger.error(f"Unexpected response when checking file existence. Status code: {response.status_code}")
                logger.error(f"Response content: {response.text}")
                return False

            # Get download URL
            download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_file_masterdata_path}:/content"
            response = requests.get(download_url, headers=headers, stream=True)

            if response.status_code == 200:
                # Create directory if it doesn't exist
                os.makedirs(folder_path, exist_ok=True)
                
                # Write file to local directory
                with open(local_file_path, 'wb') as file:
                    for chunk in response.iter_content(chunk_size=1024 * 1024):  # 1 MB chunk
                        if chunk:
                            file.write(chunk)

                logger.info(f"File downloaded successfully: {local_file_path}")
                return True
            else:
                raise Exception(f"Download failed: {response.text}")

        except Exception as e:
            logger.error(f"Error during download (attempt {attempt + 1}): {str(e)}")
            if attempt == max_retries - 1:
                logger.error("Max retries reached. Download failed.")
                return False
            else:
                time.sleep(5)  # Wait before retrying

    return False

def update_file_metadata(file_path, sharepoint_folder_path):
    try:
        file_name = os.path.basename(file_path)
        drive_id = "b!j5B2Hm8LOES-gp19kZLFNy_qoT96WGBHixaNdzAdCqJpL-7oslv5RI-IsvDE2lkX"
        update_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}"
        headers = {
            'Authorization': f'Bearer {get_current_token()}',
            'Content-Type': 'application/json'
        }
        
        # Lấy ETag hiện tại
        response = requests.get(update_url, headers=headers)
        logger.info(f"Get file info response status code: {response.status_code}")
        logger.info(f"Get file info response content: {response.text}")
        
        if response.status_code == 200:
            etag = response.headers.get('ETag')
        else:
            logger.error(f"Failed to get file info. Status code: {response.status_code}")
            return False
        
        body = {
            "fileSystemInfo": {
                "lastModifiedDateTime": datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
            }
        }
        response = requests.patch(update_url, headers=headers, json=body)
        logger.info(f"Update metadata response status code: {response.status_code}")
        logger.info(f"Update metadata response content: {response.text}")
        
        if response.status_code == 200:
            logger.info(f"Metadata updated for file: {file_name}")
            return True
        else:
            logger.error(f"Failed to update metadata for file: {file_name}")
            logger.error(f"Status code: {response.status_code}")
            logger.error(f"Response content: {response.text}")
            return False
    except Exception as e:
        logger.error(f"Error updating metadata: {str(e)}")
        return False

def update_file_version(file_path, sharepoint_folder_path):
    file_name = os.path.basename(file_path)
    drive_id = "b!j5B2Hm8LOES-gp19kZLFNy_qoT96WGBHixaNdzAdCqJpL-7oslv5RI-IsvDE2lkX"
    
    # Get item ID of the file
    check_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}"
    headers = {
        'Authorization': f'Bearer {get_current_token()}',
        'Content-Type': 'application/json'
    }
    response = requests.get(check_url, headers=headers)
    if response.status_code != 200:
        logger.error(f"Unable to get file info: {response.status_code}, {response.text}")
        return False
    
    item_id = response.json().get('id')
    
    # Check out file
    checkout_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/checkout"
    checkout_response = requests.post(checkout_url, headers=headers)
    if checkout_response.status_code != 204:
        logger.error(f"Unable to check out file: {checkout_response.status_code}, {checkout_response.text}")
        return False
    
    logger.info(f"File '{file_name}' has been checked out successfully.")
    
    # Create new version of file (check in)
    checkin_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/checkin"
    checkin_data = {
        "comment": "Updated file version",
        "checkInAs": "published"
    }
    checkin_response = requests.post(checkin_url, headers=headers, json=checkin_data)
    
    if checkin_response.status_code == 204:
        logger.info(f"New version created successfully for file '{file_name}'.")
        return True
    else:
        logger.error(f"Unable to create new version for file: {checkin_response.status_code}, {checkin_response.text}")
        return False



def combine_excel_files(folder_path, output_path):
    """
    Combines Excel files, handles date formatting, and adds master data columns
    """
    def rename_columns(df):
        """
        Rename columns according to the specified mapping
        """
        column_mapping = {
            'Document Date': 'Period',
            'Posting Date': 'Fiscal Year',
            'SKF': 'Statistical key figure',
            'Quantity': 'Statistical quantity',
            'Cost_Center': 'Cost Center'
        }
        
        return df.rename(columns=column_mapping)

    def remove_columns(df):
        """
        Remove specified columns from the dataframe
        """
        # Get list of columns to keep (exclude columns A, D, G by their index)
        columns_to_drop = df.columns[[0, 3, 6]]  # A=0, D=3, G=6
        return df.drop(columns=columns_to_drop)

    def format_period_month(date_str):
        """
        Extract month from date string and format as '0XX'
        Example: '31012025' -> '001'
        """
        try:
            # Extract month (positions 2-4)
            month = date_str[2:4]
            # Convert to number and format with leading zeros (3 digits)
            return f"{int(month):03d}"
        except Exception as e:
            logger.error(f"Error formatting period month: {str(e)}")
            return date_str

    def format_fiscal_year(date_str):
        """
        Extract year from date string (last 4 characters)
        Example: '31012025' -> '2025'
        """
        try:
            # Extract last 4 characters
            year = date_str[-4:]
            return str(year)
        except Exception as e:
            logger.error(f"Error formatting fiscal year: {str(e)}")
            return date_str

    def add_master_columns(df):
        """
        Add master data columns with default None values
        """
        new_columns = [
            "Client",
            "Client Code",
            "Business Line",
            "BU",
            "Business Model", 
            "Local 1",
            "General Name"
        ]
        
        for col in new_columns:
            df[col] = None
    
        return df
    all_dfs = []
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls')) and f not in ["Generate.xlsx", "Profit_Center_Master.xlsx", "Profit_Center_Master_temp.xlsx"]]
    
    if not excel_files:
        logger.warning("No Excel files found - creating empty Excel file with headers")
        # Create empty DataFrame with required columns
        empty_df = pd.DataFrame({
            'Controlling Area': ['VN99'],
            'Document Date': ['01'],
            'Posting Date': ['2025'],
            'Document Header Text': ['VN82'],
            'SKF': ['CRL100'],
            'Quantity': [0.0],
            'Item_Text': ['Sample'],
            'Cost_Center': ['8210001113'],
        })

        # Rename columns according to the mapping
        empty_df = rename_columns(empty_df)

        # Remove specified columns  
        empty_df = remove_columns(empty_df)

        # Save empty DataFrame with headers
        empty_df.to_excel(output_path, index=False)
        logger.info(f"Created empty Excel file with headers at {output_path}")
        return empty_df

    try:
        # Read and combine all Excel files
        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            logger.info(f"Reading file: {file}")
            
            try:
                # Read Excel file
                df = pd.read_excel(
                    file_path,
                    sheet_name=0,
                    dtype={
                        'Document Date': str,
                        'Posting Date': str,
                        'Quantity': float
                    }
                )
                
                # Forward fill Period and Fiscal Year within each file
                df['Document Date'] = df['Document Date'].ffill()
                df['Posting Date'] = df['Posting Date'].ffill()

                # Format Period (month) and Fiscal Year
                df['Document Date'] = df['Document Date'].apply(format_period_month)
                df['Posting Date'] = df['Posting Date'].apply(format_fiscal_year)
                
                # Convert other columns to string
                for col in df.columns:
                    if col not in ['Document Date', 'Posting Date', 'Quantity']:
                        df[col] = df[col].astype(str)
                
                all_dfs.append(df)
                logger.info(f"Successfully read and formatted {file}")
                
            except Exception as e:
                logger.error(f"Error processing file {file}: {str(e)}")
                continue
        
        if not all_dfs:
            logger.error("No valid data found in any Excel files")
            return None
            
        # Combine all dataframes
        combined_df = pd.concat(all_dfs, ignore_index=True)
        
        # Rename columns
        combined_df = rename_columns(combined_df)
        
        # Remove specified columns
        combined_df = remove_columns(combined_df)
        
        # Add master data columns
        combined_df = add_master_columns(combined_df)
        
        # Save to Excel
        combined_df.to_excel(output_path, index=False)
        logger.info(f"Successfully combined {len(excel_files)} files into {output_path}")
        return combined_df
            
    except Exception as e:
        logger.error(f"Error combining Excel files: {str(e)}")
        if os.path.exists(output_path):
            os.remove(output_path)
            logger.info(f"Removed incomplete output file: {output_path}")
        raise
    
    return None

def download_file_wrapper(args: Tuple[str, str, str]) -> Tuple[bool, str, str]:
    id, file_name, folder_path = args
    sharepoint_folder_path = f"/Orchestrator/.Portal/SKF/Generate/{id}/{file_name}"
    success = download_file_from_sharepoint(sharepoint_folder_path, folder_path)
    return success, sharepoint_folder_path, id

def extract_cr58d_filename(document_str):
    try:
        # Find the start and end indices of cr58d_filename
        start_index = document_str.find("'cr58d_filename': '") + len("'cr58d_filename': '")
        end_index = document_str.find("'", start_index)

        # Extract and return the cr58d_filename value
        return document_str[start_index:end_index]
    except Exception as e:
        logger.error(f"Error extracting cr58d_filename: {e}")
        return None
