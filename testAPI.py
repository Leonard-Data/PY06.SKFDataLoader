import requests
import os
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from APIs.requests import RequestAPI
from auth.graphAPI import get_current_token
from utils.functions import *

folder_path = create_folder()

combined_file_path = os.path.join(folder_path, "Generate.xlsx")
combine_excel_files(folder_path, combined_file_path)
