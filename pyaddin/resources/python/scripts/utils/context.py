'''Manage parameters context over modules.'''

import os
import logging
import win32com.client


KEY_CALLER = 'caller'
KEY_LOGGER = 'logger'


CONFIG_FILENAME = 'main.cfg'



def start():
    global _dict
    _dict = {}

def set(key, value):
    _dict[key] = value

def get(key, default_value=None):
    if key in _dict:
        return _dict[key]
    else:
        return default_value


def set_caller(caller_name:str):
    '''set workbook instance calling this script.'''
    wb = __get_caller_workbook(caller_name)
    set(KEY_CALLER, wb)


def get_caller() -> 'win32com.WorkBook':
    '''get workbook instance calling this script.'''
    return get(KEY_CALLER)


def set_logger(working_path:str):
    '''set logger based on configuration file under working path.'''
    logger = __config_logger(working_path)
    set(KEY_LOGGER, logger)


def get_logger() -> logging.Logger:
    '''get logger.'''
    return get(KEY_LOGGER)


def __get_caller_workbook(name:str):
    '''Get Workbook instance (win32com) by name.

    Args:
        name (str): Workbook name.
    '''
    app = win32com.client.Dispatch('Excel.Application')
    for wb in app.Workbooks:
        if wb.Name == name: return wb
    return None


def __config_logger(working_path:str) -> logging.Logger:
    '''Log settings based on configuration file under working path.'''
    # log files
    config_file = os.path.join(working_path, CONFIG_FILENAME)
    output_file, error_file, log_file = __check_log_files(config_file=config_file)

    # full level log
    log = logging.FileHandler(filename=log_file, mode='w', encoding='utf-8')
    log.setLevel(level=logging.DEBUG)
    fmt = logging.Formatter(fmt="%(asctime)s - %(name)s - %(levelname)s -%(module)s:  %(message)s", datefmt='%Y-%m-%d %H:%M:%S')
    log.setFormatter(fmt)

    # normal output
    output = logging.FileHandler(filename=output_file, mode='w', encoding='utf-8')
    output.setLevel(level=logging.INFO)

    # error
    error = logging.FileHandler(filename=error_file, mode='w', encoding='utf-8')
    error.setLevel(level=logging.ERROR)

    # console
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)

    # logger
    logger = logging.Logger(name='addin_logger', level=logging.DEBUG)
    logger.addHandler(log)
    logger.addHandler(output)
    logger.addHandler(error)
    logger.addHandler(console)

    return logger


def __check_log_files(config_file:str):
    '''Get log file names from configuration file.'''
    # default value
    output_path = 'outputs'
    output_name, error_name, log_name = 'output.log', 'errors.log', 'log.log'

    # check config file
    with open(config_file, 'r') as f:
        while True:
            line = f.readline()
            if not line:
                break
            elif line.startswith('[output]'):
                output_path = f.readline().strip()
            elif line.startswith('[stdout]'):
                output_name = f.readline().strip()
            elif line.startswith('[stderr]'):
                error_name = f.readline().strip()
            elif line.startswith('[log]'):
                log_name = f.readline().strip()
    # create output path
    if not os.path.exists(output_path):
        os.mkdir(output_path)
    
    # join path
    output_file = os.path.join(output_path, output_name)
    error_file = os.path.join(output_path, error_name)
    log_file = os.path.join(output_path, log_name)
    
    return output_file, error_file, log_file