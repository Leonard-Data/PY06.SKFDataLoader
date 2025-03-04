"""
BY: PHUNG H.BINH 
PROJECT MADE WITH: Qt Designer, PySide6 and GraphAPI technical
Purpose: Reporting team dashboard 
Version: 1.0.1

descriptions:
The logger class provides a simple tool for tracing function execution behavior.

"""
import logging
import os
import time

def setup_logger(module):
    """
    setup file loggers
    """
    logger = logging.getLogger(module)
    logger.setLevel(logging.DEBUG)  # Set the minimum log level
    _global_path = os.path.join(os.path.expanduser("~"), '.logs/PY06')
    if not os.path.isdir(_global_path):
        os.makedirs(_global_path, exist_ok=True)
    log_name = 'Execution_'+time.strftime("%Y%m%d")
    if not os.path.isdir(f"{_global_path}"):
        os.makedirs(f"{_global_path}", exist_ok=True)
    file_handler = logging.FileHandler(f"{_global_path}/{log_name}.log", mode='a')
    file_handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s %(levelname)s [%(name)s] %(message)s',datefmt="%Y-%m-%d %H:%M:%S")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger

def clear_logging():
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
