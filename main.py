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
from utils.exceptions import (
    SKFError, 
    NoDataFoundError, 
    FileNotFoundError, 
    SharePointError,
    TokenNotAvailableError,
    DataError,
    ProcessError,
    get_error_code
)
from utils.retry import retry_with_backoff, retry_heavy
from auth.graphAPI import stop_monitoring, get_current_token, monitor_token
from auth.dataverseAPI import stop_monitoring_dataverse, get_current_token_dataverse, monitor_token_dataverse

# Setup stdout encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Setup global logger at module level
logger = setup_logger(os.path.basename(__file__))

def check_running_instances():
    """Check if another instance is already running"""
    import psutil
    current_pid = os.getpid()
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if (proc.info['name'] == 'python.exe' and 
                proc.info['pid'] != current_pid and
                'main.py' in ' '.join(proc.info['cmdline'] or [])):
                logger.warning(f"Another instance is running (PID: {proc.info['pid']})")
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

@retry_with_backoff(max_retries=1, timeout=120.0)
def download_masterdata_with_retry(sharepoint_masterdata_file_path: str, folder_path: str) -> bool:
    """Download masterdata with automatic retry"""
    return download_file_from_sharepoint(sharepoint_masterdata_file_path, folder_path)

@retry_with_backoff(max_retries=1, timeout=60.0)  
def get_api_data_with_retry() -> list:
    """Get API data with automatic retry"""
    get_data_request = RequestAPI()
    return get_data_request.get_codes()

@retry_heavy()
def upload_to_sharepoint_with_retry(output_path: str, sharepoint_folder_path: str) -> bool:
    """Upload file to SharePoint with automatic retry"""
    return upload_file_to_sharepoint(output_path, sharepoint_folder_path)

def log_performance_metrics(start_time: float, operation: str = ""):
    """Log performance metrics and warn if slow"""
    execution_time = time.time() - start_time
    if execution_time > 600:  # 10 minutes
        logger.warning(f"⚠️ Slow execution ({operation}): {execution_time:.2f}s")
    elif execution_time > 900:  # 15 minutes  
        logger.error(f"🚨 Very slow execution ({operation}): {execution_time:.2f}s")

def main(year: int, month: int):
    start_time = time.time()
    logger.info("="*50)
    logger.info("STARTING MAIN PROCESS")
    logger.info("="*50)
    
    try:
        # Kill existing processes with better error handling
        logger.info("Killing existing processes...")
        try:
            kill_process(["Excel.exe", "saplogon.exe"])
        except ProcessError as e:
            logger.warning(f"Process cleanup warning: {e}")
            # Continue execution as this is not critical

        # Create/clean folder
        logger.info("Creating/cleaning folder...")
        folder_path = create_folder()
        logger.info(f"Folder path: {folder_path}")
        
        # Setup paths
        sharepoint_masterdata_file_path = f"/General/11.%20Master%20Data/Profit_Center_Master.xlsx"
        sharepoint_folder_path = f"/Orchestrator/.Portal/SKF/Sources/{year}"
        report_file_name = f"ACP.CM_{year}.xlsx"
        
        # Get token
        logger.info("Getting access token...")
        access_token = get_current_token()
        if not access_token:
            raise TokenNotAvailableError("Access token not available")
        logger.info("[SUCCESS] Access token available")
        logger.info(f"Report file name: {report_file_name}")
        
        # Kill SAP processes
        try:
            kill_process(["sapgui.exe", "saplogon.exe"])
        except ProcessError as e:
            logger.warning(f"SAP process cleanup warning: {e}")

        logger.info("Starting the main process...")
            
        # Download masterdata with automatic retry
        logger.info("Downloading masterdata from SharePoint...")
        logger.info(f"SharePoint path: {sharepoint_masterdata_file_path}")
        
        masterdata_downloaded = download_masterdata_with_retry(sharepoint_masterdata_file_path, folder_path)
        
        if not masterdata_downloaded:
            raise SharePointError("Failed to download masterdata file")
        
        logger.info("[SUCCESS] MasterData download completed")
        
        # Process masterdata
        full_masterdata_file_path = os.path.join(folder_path, "Profit_Center_Master.xlsx")
        logger.info(f"Processing MasterData at: {full_masterdata_file_path}")
        
        processed_MasterData_file = ProcessMasterData(full_masterdata_file_path)
        if "Error" in processed_MasterData_file:
            raise DataError(f"MasterData processing failed: {processed_MasterData_file}")
        
        logger.info("[SUCCESS] Processing MasterData completed successfully!")
                
        # Get data request with automatic retry
        logger.info("Getting data request from API...")
        data_request = get_api_data_with_retry()
        
        if not data_request:
            raise NoDataFoundError("No data received from API request")
        
        logger.info(f"[SUCCESS] Data request successful. Retrieved {len(data_request)} records")
        
        # Process data
        df = pd.DataFrame([data.dict() for data in data_request])
        logger.info(f"DataFrame created with {len(df)} rows")

        # Convert and filter data
        df['status_request'] = pd.to_numeric(df['status_request'], errors='coerce')
        df['year'] = pd.to_numeric(df['year'], errors='coerce')
        filtered_df = df[(df['status_request'].isin([1,2,5])) & (df['period'] == month) & (df['year'] == year)]
        logger.info(f"Filtered DataFrame: {len(filtered_df)} rows for year {year}, month {month}")
        
        if len(filtered_df) == 0:
            logger.warning(f"No data found for year {year}, month {month} - creating empty file")
            # Create empty combined file and continue process
            combined_file_path = os.path.join(folder_path, report_file_name)
            combine_excel_files(folder_path, combined_file_path)  # This will create empty file
            logger.info(f"[SUCCESS] Empty file created: {combined_file_path}")
            successful_downloads = 0
        else:
            # Extract filenames and create download dictionary
            filtered_df = filtered_df.copy()  # Fix SettingWithCopyWarning
            filtered_df['cr58d_filename'] = filtered_df['cr58d_GenFileID'].apply(extract_cr58d_filename)

            id_file_dict = {row['id']: row['cr58d_filename'] for index, row in filtered_df.iterrows() if row['cr58d_filename']}
            
            total_files = len(id_file_dict)
            logger.info(f"Found {total_files} files to download")
            
            if total_files == 0:
                logger.warning("No valid files found to download - creating empty file")
                # Create empty combined file and continue process
                combined_file_path = os.path.join(folder_path, report_file_name)
                combine_excel_files(folder_path, combined_file_path)  # This will create empty file
                logger.info(f"[SUCCESS] Empty file created: {combined_file_path}")
                successful_downloads = 0
            else:
                # Download files with parallel processing
                logger.info(f"Starting parallel download of {total_files} files...")
                download_start = time.time()
                
                successful_downloads = 0
                download_tasks = [(id, file_name, folder_path) for id, file_name in id_file_dict.items()]
                
                # Use ThreadPoolExecutor with reasonable max_workers for 15-min schedule
                with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:  # Reduced from 5 to 3
                    futures = [executor.submit(download_file_wrapper, task) for task in download_tasks]
                    
                    for future in concurrent.futures.as_completed(futures):
                        try:
                            success, path, id = future.result()
                            if success:
                                successful_downloads += 1
                                logger.info(f"[SUCCESS] File downloaded: {path} ({successful_downloads}/{total_files})")
                            else:
                                logger.error(f"[FAILED] Download failed: {path}")
                        except Exception as e:
                            logger.error(f"[FAILED] Download task failed: {e}")

                log_performance_metrics(download_start, f"Download {total_files} files")
                logger.info(f"Download summary: {successful_downloads}/{total_files} files successfully downloaded")

                if successful_downloads == 0:
                    logger.warning("No files were downloaded successfully - creating empty file")
                    # Create empty combined file and continue process
                    combined_file_path = os.path.join(folder_path, report_file_name)
                    combine_excel_files(folder_path, combined_file_path)  # This will create empty file
                    logger.info(f"[SUCCESS] Empty file created: {combined_file_path}")
                else:
                    # Combine Excel files (normal case)
                    logger.info("Combining Excel files...")
                    combined_file_path = os.path.join(folder_path, report_file_name)
                    
                    try:
                        combine_excel_files(folder_path, combined_file_path)
                    except Exception as e:
                        raise DataError(f"Failed to combine Excel files: {e}")
                    
                    logger.info(f"[SUCCESS] Combined file saved to: {combined_file_path}")

        # Consolidate data (works for both empty and real data)
        logger.info("Starting data consolidation with masterdata...")
        output_path = Consolidation_report_masterdata(combined_file_path, processed_MasterData_file)
        
        if "Error" in output_path:
            raise DataError(f"Data consolidation failed: {output_path}")
        
        logger.info(f"[SUCCESS] Process completed successfully! Output file: {output_path}")
            
        # Verify output file
        if not os.path.exists(output_path):
            raise FileNotFoundError(f"Output file not found: {output_path}")
        
        logger.info(f"[SUCCESS] Output file verified: {output_path}")
        logger.info(f"File size: {os.path.getsize(output_path):,} bytes")
        
        # SharePoint operations with retry
        drive_id = "b!j5B2Hm8LOES-gp19kZLFNy_qoT96WGBHixaNdzAdCqJpL-7oslv5RI-IsvDE2lkX"
        
        logger.info("Checking SharePoint access...")
        if not check_sharepoint_access(access_token, drive_id):
            raise SharePointError("No access to SharePoint. Please check permissions.")
        logger.info("[SUCCESS] SharePoint access verified")
        
        logger.info("Checking SharePoint folder...")
        if not check_sharepoint_folder(access_token, drive_id, sharepoint_folder_path):
            raise SharePointError("Failed to verify SharePoint folder. Please check the path.")
        logger.info("[SUCCESS] SharePoint folder verified")
            
        # Upload with automatic retry
        logger.info(f"Starting upload to SharePoint...")
        logger.info(f"SharePoint folder path: {sharepoint_folder_path}")
        
        upload_success = upload_to_sharepoint_with_retry(output_path, sharepoint_folder_path)
        
        if not upload_success:
            raise SharePointError("Failed to upload file to SharePoint")
                        
        logger.info(f"[SUCCESS] File uploaded to SharePoint successfully!")
        
        # Update file version
        logger.info("Updating file version...")
        if update_file_version(output_path, sharepoint_folder_path):
            logger.info("[SUCCESS] File version updated successfully.")                            
        else:
            logger.warning("[WARNING] Failed to update file version.")
        
        # Final cleanup
        logger.info("Cleaning up processes...")
        try:
            kill_process(["Excel.exe", "saplogon.exe"])
        except ProcessError as e:
            logger.warning(f"Final cleanup warning: {e}")
        
    except SKFError as e:
        error_code = get_error_code(e)
        logger.error(f"[SKF ERROR] {error_code}: {e}")
        logger.error(f"Retryable: {e.retryable}")
        raise
    except Exception as e:
        logger.error(f"[UNEXPECTED ERROR] {type(e).__name__}: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        # Wrap unexpected errors
        wrapped_error = SKFError(f"Unexpected error: {e}", retryable=False)
        raise wrapped_error from e
    finally:
        end_time = time.time()  
        elapsed_time = end_time - start_time  
        logger.info("="*50)
        logger.info(f"PROCESS COMPLETED - Total execution time: {elapsed_time:.2f} seconds")
        logger.info("="*50)
        
        # Performance warning
        if elapsed_time > 600:  # 10 minutes
            logger.warning(f"⚠️ Process took {elapsed_time:.2f}s - consider optimization")
        
        # Final cleanup
        try:
            kill_process(["Excel.exe", "saplogon.exe"])
        except:
            pass  # Ignore cleanup errors in finally block
        
        stop_monitoring()
        stop_monitoring_dataverse()

if __name__ == '__main__':
    # Check for existing instances
    if check_running_instances():
        logger.error("Another instance is already running. Exiting...")
        sys.exit(1)
    
    # Setup console handler with shorter time format
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', 
                                datefmt='%y-%m-%d %H:%M')
    console_handler.setFormatter(formatter)
    
    # Add console handler to root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers.clear()
    root_logger.addHandler(console_handler)
    
    logger.info("Application starting...")
    
    # Parse arguments
    parser = argparse.ArgumentParser()
    parser.add_argument('year', type=int, help='Year to process')
    parser.add_argument('month', type=int, help='Month to process')
    args = parser.parse_args()
    
    year = args.year
    month = args.month
    
    logger.info(f"Arguments received - Year: {year}, Month: {month}")
    
    # Start monitoring threads
    logger.info("Starting monitoring threads...")
    monitor_thread_graph = threading.Thread(target=monitor_token, daemon=True)
    monitor_thread_dataverse = threading.Thread(target=monitor_token_dataverse, daemon=True)
    monitor_thread_graph.start()
    monitor_thread_dataverse.start()

    # Wait for tokens with timeout
    logger.info("Waiting for tokens...")
    token_timeout = 60  # 1 minute timeout
    start_wait = time.time()
    
    while (get_current_token() is None or get_current_token_dataverse() is None):
        if time.time() - start_wait > token_timeout:
            logger.error("Token acquisition timeout!")
            sys.exit(1)
        time.sleep(1)
    
    logger.info("[SUCCESS] All tokens obtained")

    try:
        main(year, month)
        logger.info("[SUCCESS] Main execution completed successfully")
    except SKFError as e:
        logger.error(f"[FAILED] Main execution failed: [{get_error_code(e)}] {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"[FAILED] Unexpected error: {e}")
        sys.exit(1)
    finally:
        logger.info("Stopping monitoring threads...")
        stop_monitoring()
        stop_monitoring_dataverse()
        # Give threads time to stop gracefully
        time.sleep(2)
        logger.info("[COMPLETED] Application terminated")