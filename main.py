import argparse
import logging
import os
import threading
import time
import sys
import configparser  # Import library to read configuration
from logging import config
from APIs.requests import RequestAPI
from SAP.extend import ExtendFunctions
from utils.functions import *
from utils.log import setup_logger
from data.get_credential_data import get_credential_data
from auth.graphAPI import stop_monitoring, get_current_token, monitor_token
from auth.dataverseAPI import stop_monitoring_dataverse, get_current_token_dataverse, monitor_token_dataverse
from requests.exceptions import RequestException
import concurrent.futures
from typing import Tuple

sys.stdout.reconfigure(encoding='utf-8')

logger = logging.getLogger(__name__)

def main(year: int, month: int):
    start_time = time.time()
    try:
        kill_process(["Excel.exe", "saplogon.exe"])

        # Create a folder to store downloaded files - If the folder exists, delete the files inside
        folder_path = create_folder()
        
        #setup sharepoint masterdata file path
        sharepoint_masterdata_file_path = f"/General/11.%20Master%20Data/Profit_Center_Master.xlsx"
        
        #setup sharepoint folder path
        sharepoint_folder_path = f"/Orchestrator/.Portal/SKF/Sources/{year}"
        
        # Get token and set URL
        access_token = get_current_token()
        
        #setup file name
        report_file_name = f"ACP.CM_{year}.xlsx"
        
        # Setup logger
        logger = setup_logger(os.path.basename(__file__))

        # Kill SAP processes if any
        kill_process(["sapgui.exe", "saplogon.exe"])

        logger.info("Starting the process...")
            
        masterdata_downloaded = download_file_from_sharepoint(sharepoint_masterdata_file_path, folder_path)
        logger.info(f"MasterData download status: {masterdata_downloaded}")
        full_masterdata_file_path = os.path.join(folder_path, "Profit_Center_Master.xlsx")
        logger.info("Processing MasterData...")
        processed_MasterData_file = ProcessMasterData(full_masterdata_file_path)
        
        if "Error" in processed_MasterData_file:
            logger.error(f"Error in processing MasterData: {processed_MasterData_file}")
            raise Exception(processed_MasterData_file)
        logger.info("Processing MasterData completed successfully!")
                
        get_data_request = RequestAPI()
        data_request = get_data_request.get_codes()
        
        if data_request:
            logger.info("Data request successful.")
            # Convert data_request to a pandas DataFrame
            df = pd.DataFrame([data.dict() for data in data_request])

            # Convert the status_request column to numeric
            df['status_request'] = pd.to_numeric(df['status_request'], errors='coerce')

            # Filter the DataFrame for status_request = 5 and period = month
            filtered_df = df[(df['status_request'].isin([1,2,5])) & (df['period'] == month)]
            
            # Extract cr58d_filename from cr58d_GenFileID
            filtered_df['cr58d_filename'] = filtered_df['cr58d_GenFileID'].apply(extract_cr58d_filename)
            
            # Save the filtered DataFrame to an Excel file
            # filtered_df.to_excel("filtered_data_request.xlsx", index=False)

            # Create a dictionary for ID and file name pairs
            id_file_dict = {row['id']: row['cr58d_filename'] for index, row in filtered_df.iterrows() if row['cr58d_filename']}
            
            total_files = len(id_file_dict)
            successful_downloads = 0
            logger.info(f"Starting to download {total_files} files in parallel...")
            
            # Create download tasks
            download_tasks = [(id, file_name, folder_path) 
                            for id, file_name in id_file_dict.items()]
            
            # Use ThreadPoolExecutor to download files in parallel
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(download_file_wrapper, task) 
                          for task in download_tasks]
                
                for future in concurrent.futures.as_completed(futures):
                    success, path, id = future.result()
                    if success:
                        successful_downloads += 1
                        logger.info(f"File downloaded successfully from {path} ({successful_downloads}/{total_files})")
                    else:
                        logger.error(f"Failed to download file from {path} ({successful_downloads}/{total_files})")

            logger.info(f"Download completed: {successful_downloads}/{total_files} files successfully downloaded")

            # Combine all Excel files in the folder into one file
            combined_file_path = os.path.join(folder_path, report_file_name)
            combine_excel_files(folder_path, combined_file_path)
            logger.info(f"Combined file saved to: {combined_file_path}")
            output_path = Consolidation_report_masterdata(combined_file_path, processed_MasterData_file)
            
            if "Error" in output_path:
                    logger.error(f"Error in consolidating data: {output_path}")
                    raise Exception(output_path)
            logger.info(f"Process completed successfully! Output file: {output_path}")
                
            drive_id = "b!j5B2Hm8LOES-gp19kZLFNy_qoT96WGBHixaNdzAdCqJpL-7oslv5RI-IsvDE2lkX"
                
            if os.path.exists(output_path):
                print(f"File found: {output_path}")
                if not check_sharepoint_access(access_token, drive_id):
                    print("No access to SharePoint. Please check permissions.")
                    return
                if not check_sharepoint_folder(access_token, drive_id, sharepoint_folder_path):
                    print("Failed to verify SharePoint folder. Please check the path.")
                    return
                try:
                    upload_success = upload_file_to_sharepoint(output_path, sharepoint_folder_path)
                    logger.info(f"Starting upload of file: {output_path}")
                    logger.info(f"File size: {os.path.getsize(output_path)} bytes")
                    logger.info(f"SharePoint folder path: {sharepoint_folder_path}")
                    if upload_success:                            
                        logger.info(f"File {output_path} uploaded to SharePoint successfully.")
                        if update_file_version(output_path, sharepoint_folder_path):
                            print("File version updated successfully.")                            
                        else:
                            print("Failed to update file version.")
                    else:
                        print("Failed to upload file to SharePoint.")
                        logger.error(f"Failed to upload file {output_path} to SharePoint.")
                except requests.exceptions.RequestException as e:
                    print(f"Network error occurred: {e}")
                    logger.error(f"Network error during file upload: {e}")
                except Exception as e:
                    print(f"An unexpected error occurred during file upload: {e}")
                    logger.error(f"Unexpected error during file upload: {e}")
            else:
                print(f"Error: File not found at {output_path}")
            
            kill_process(["Excel.exe", "saplogon.exe"])
            stop_monitoring()
            stop_monitoring_dataverse()
        else:
            logger.error("Data request returned no data or encountered an error.")
        
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")
        kill_process(["Excel.exe", "saplogon.exe"])
        stop_monitoring()
        stop_monitoring_dataverse()
    finally:
        end_time = time.time()  
        elapsed_time = end_time - start_time  
        logger.info(f"Total execution time: {elapsed_time:.2f} seconds")
        stop_monitoring()
        stop_monitoring_dataverse()


if __name__ == '__main__':
    
    parser = argparse.ArgumentParser()
    parser.add_argument('year', type=int, help='Year to process')
    parser.add_argument('month', type=int, help='Month to process')
    args = parser.parse_args()
    
    year = args.year
    month = args.month
    
    monitor_thread_graph = threading.Thread(target=monitor_token)
    monitor_thread_dataverse = threading.Thread(target=monitor_token_dataverse)
    monitor_thread_graph.start()
    monitor_thread_dataverse.start()

    while get_current_token() is None:
        time.sleep(1)
    while get_current_token_dataverse() is None:
        time.sleep(1)

    main(year,month)

    stop_monitoring()
    monitor_thread_graph.join()
    stop_monitoring_dataverse()
    monitor_thread_dataverse.join()
