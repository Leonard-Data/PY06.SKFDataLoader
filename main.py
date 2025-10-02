import argparse
import logging
import os
import threading
import time
import sys
import concurrent.futures
import io

from APIs.requests import RequestAPI
from utils.functions import *
from utils.log import setup_logger
from auth.graphAPI import stop_monitoring, get_current_token, monitor_token
from auth.dataverseAPI import stop_monitoring_dataverse, get_current_token_dataverse, monitor_token_dataverse

# Setup stdout encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Setup global logger at module level
logger = setup_logger(os.path.basename(__file__))

def main(year: int, month: int):
    start_time = time.time()
    logger.info("="*50)
    logger.info("STARTING MAIN PROCESS")
    logger.info("="*50)
    
    try:
        logger.info("Killing existing processes...")
        kill_process(["Excel.exe", "saplogon.exe"])

        # Create a folder to store downloaded files - If the folder exists, delete the files inside
        logger.info("Creating/cleaning folder...")
        folder_path = create_folder()
        logger.info(f"Folder path: {folder_path}")
        
        #setup sharepoint masterdata file path
        sharepoint_masterdata_file_path = f"/General/11.%20Master%20Data/Profit_Center_Master.xlsx"
        
        #setup sharepoint folder path
        sharepoint_folder_path = f"/Orchestrator/.Portal/SKF/Sources/{year}"
        
        # Get token and set URL
        logger.info("Getting access token...")
        access_token = get_current_token()
        logger.info(f"Access token status: {'[SUCCESS] Available' if access_token else '[FAILED] Not available'}")
        
        #setup file name
        report_file_name = f"ACP.CM_{year}.xlsx"
        logger.info(f"Report file name: {report_file_name}")
        
        # Kill SAP processes if any
        kill_process(["sapgui.exe", "saplogon.exe"])

        logger.info("Starting the main process...")
            
        # Download masterdata with retry mechanism
        logger.info("Downloading masterdata from SharePoint...")
        logger.info(f"SharePoint path: {sharepoint_masterdata_file_path}")
        
        # Try download with better error handling
        masterdata_downloaded = False
        max_download_attempts = 1
        
        for attempt in range(max_download_attempts):
            try:
                logger.info(f"Download attempt {attempt + 1}/{max_download_attempts}")
                masterdata_downloaded = download_file_from_sharepoint(sharepoint_masterdata_file_path, folder_path)
                
                if masterdata_downloaded:
                    logger.info("[SUCCESS] MasterData download completed")
                    break
                else:
                    logger.warning(f"Download attempt {attempt + 1} failed")
                    if attempt < max_download_attempts - 1:
                        wait_time = (attempt + 1) * 30  # Increasing wait time
                        logger.info(f"Waiting {wait_time} seconds before next attempt...")
                        time.sleep(wait_time)
                        
            except Exception as e:
                logger.error(f"Download attempt {attempt + 1} failed with error: {e}")
                if attempt == max_download_attempts - 1:
                    raise
                time.sleep(30)
        
        if not masterdata_downloaded:
            raise Exception("Failed to download masterdata file after all attempts")
        
        full_masterdata_file_path = os.path.join(folder_path, "Profit_Center_Master.xlsx")
        logger.info(f"Processing MasterData at: {full_masterdata_file_path}")
        processed_MasterData_file = ProcessMasterData(full_masterdata_file_path)
        
        if "Error" in processed_MasterData_file:
            logger.error(f"Error in processing MasterData: {processed_MasterData_file}")
            raise Exception(processed_MasterData_file)
        logger.info("[SUCCESS] Processing MasterData completed successfully!")
                
        # Get data request with automatic retry
        logger.info("Getting data request from API...")
        get_data_request = RequestAPI()
        data_request = get_data_request.get_codes()
        
        if data_request:
            logger.info(f"[SUCCESS] Data request successful. Retrieved {len(data_request)} records")
            
            # Convert data_request to a pandas DataFrame
            df = pd.DataFrame([data.dict() for data in data_request])
            logger.info(f"DataFrame created with {len(df)} rows")

            # Convert the status_request column to numeric
            df['status_request'] = pd.to_numeric(df['status_request'], errors='coerce')

            # Filter the DataFrame for status_request = 5 and period = month
            filtered_df = df[(df['status_request'].isin([1,2,5])) & (df['period'] == month)]
            logger.info(f"Filtered DataFrame: {len(filtered_df)} rows for year {year}, month {month}")
            
            # Extract cr58d_filename from cr58d_GenFileID
            filtered_df['cr58d_filename'] = filtered_df['cr58d_GenFileID'].apply(extract_cr58d_filename)

            # Create a dictionary for ID and file name pairs
            id_file_dict = {row['id']: row['cr58d_filename'] for index, row in filtered_df.iterrows() if row['cr58d_filename']}
            
            total_files = len(id_file_dict)
            logger.info(f"Found {total_files} files to download")
            
            if total_files == 0:
                logger.warning("No files to download!")
                return
            
            successful_downloads = 0
            logger.info(f"Starting parallel download of {total_files} files...")
            
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
                        logger.info(f"[SUCCESS] File downloaded: {path} ({successful_downloads}/{total_files})")
                    else:
                        logger.error(f"[FAILED] Download failed: {path} ({successful_downloads}/{total_files})")

            logger.info(f"Download summary: {successful_downloads}/{total_files} files successfully downloaded")

            if successful_downloads == 0:
                raise Exception("No files were downloaded successfully")

            # Combine all Excel files in the folder into one file
            logger.info("Combining Excel files...")
            combined_file_path = os.path.join(folder_path, report_file_name)
            combine_excel_files(folder_path, combined_file_path)
            logger.info(f"[SUCCESS] Combined file saved to: {combined_file_path}")
            
            logger.info("Starting data consolidation with masterdata...")
            output_path = Consolidation_report_masterdata(combined_file_path, processed_MasterData_file)
            
            if "Error" in output_path:
                logger.error(f"Error in consolidating data: {output_path}")
                raise Exception(output_path)
            logger.info(f"[SUCCESS] Process completed successfully! Output file: {output_path}")
                
            drive_id = "b!j5B2Hm8LOES-gp19kZLFNy_qoT96WGBHixaNdzAdCqJpL-7oslv5RI-IsvDE2lkX"
                
            if os.path.exists(output_path):
                logger.info(f"[SUCCESS] Output file verified: {output_path}")
                logger.info(f"File size: {os.path.getsize(output_path):,} bytes")
                
                # Check SharePoint access and folder with retry
                logger.info("Checking SharePoint access...")
                if not check_sharepoint_access(access_token, drive_id):
                    logger.error("[FAILED] No access to SharePoint. Please check permissions.")
                    return
                logger.info("[SUCCESS] SharePoint access verified")
                
                logger.info("Checking SharePoint folder...")
                if not check_sharepoint_folder(access_token, drive_id, sharepoint_folder_path):
                    logger.error("[FAILED] Failed to verify SharePoint folder. Please check the path.")
                    return
                logger.info("[SUCCESS] SharePoint folder verified")
                    
                try:
                    logger.info(f"Starting upload to SharePoint...")
                    logger.info(f"SharePoint folder path: {sharepoint_folder_path}")
                    
                    # Upload with retry mechanism
                    upload_success = upload_file_to_sharepoint(output_path, sharepoint_folder_path)
                    
                    if upload_success:                            
                        logger.info(f"[SUCCESS] File uploaded to SharePoint successfully!")
                        logger.info("Updating file version...")
                        if update_file_version(output_path, sharepoint_folder_path):
                            logger.info("[SUCCESS] File version updated successfully.")                            
                        else:
                            logger.warning("[WARNING] Failed to update file version.")
                    else:
                        logger.error("[FAILED] Failed to upload file to SharePoint.")
                        raise Exception("Upload failed")
                        
                except Exception as e:
                    logger.error(f"Error during file upload: {e}")
                    raise
            else:
                logger.error(f"[FAILED] Output file not found: {output_path}")
                raise FileNotFoundError(f"Output file not found: {output_path}")
            
            logger.info("Cleaning up processes...")
            kill_process(["Excel.exe", "saplogon.exe"])
            
        else:
            logger.error("[FAILED] Data request returned no data or encountered an error.")
            raise Exception("No data received from API request")
        
    except Exception as e:
        logger.error(f"[ERROR] An unexpected error occurred: {e}")
        logger.error(f"Error type: {type(e).__name__}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise
    finally:
        end_time = time.time()  
        elapsed_time = end_time - start_time  
        logger.info("="*50)
        logger.info(f"PROCESS COMPLETED - Total execution time: {elapsed_time:.2f} seconds")
        logger.info("="*50)
        kill_process(["Excel.exe", "saplogon.exe"])
        stop_monitoring()
        stop_monitoring_dataverse()

if __name__ == '__main__':
    # Setup console handler with shorter time format
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    # Custom formatter with shorter time format (YY-MM-DD HH:MM)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', 
                                datefmt='%y-%m-%d %H:%M')
    console_handler.setFormatter(formatter)
    
    # Add console handler to root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    # Clear existing handlers to avoid duplicates
    root_logger.handlers.clear()
    root_logger.addHandler(console_handler)
    
    logger.info("Application starting...")
    
    parser = argparse.ArgumentParser()
    parser.add_argument('year', type=int, help='Year to process')
    parser.add_argument('month', type=int, help='Month to process')
    args = parser.parse_args()
    
    year = args.year
    month = args.month
    
    logger.info(f"Arguments received - Year: {year}, Month: {month}")
    
    logger.info("Starting monitoring threads...")
    monitor_thread_graph = threading.Thread(target=monitor_token)
    monitor_thread_dataverse = threading.Thread(target=monitor_token_dataverse)
    monitor_thread_graph.start()
    monitor_thread_dataverse.start()

    logger.info("Waiting for tokens...")
    while get_current_token() is None:
        time.sleep(1)
    while get_current_token_dataverse() is None:
        time.sleep(1)
    logger.info("[SUCCESS] All tokens obtained")

    try:
        main(year, month)
        logger.info("[SUCCESS] Main execution completed successfully")
    except Exception as e:
        logger.error(f"[FAILED] Main execution failed: {e}")
    finally:
        logger.info("Stopping monitoring threads...")
        stop_monitoring()
        monitor_thread_graph.join()
        stop_monitoring_dataverse()
        monitor_thread_dataverse.join()
        logger.info("[COMPLETED] Application terminated")
