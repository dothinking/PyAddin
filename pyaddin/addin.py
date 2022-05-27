import os
import logging
import shutil
import win32com.client
from .xlam.ui import UI
from .xlam.vba import VBA
from .share import AddInException


# logging
logging.basicConfig(
    level=logging.INFO, 
    format="[%(levelname)s] %(message)s")


# configuration path
SCRIPT_PATH = os.path.dirname(os.path.abspath(__file__)) 
RESOURCE_PATH = os.path.join(SCRIPT_PATH, 'resources')

RESOURCE_ADDIN = 'xlam'
RESOURCE_PYTHON = 'scripts'
RESOURCE_VBA = 'vba'
PYTHON_MAIN = 'main.py'
PYTHON_CONFIG = 'main.cfg'
VBA_GENERAL = 'General'
VBA_MENU = 'Ribbon'
VBA_USER_MENU = 'UserRibbon'
CUSTOM_UI = 'CustomUI.xml'


class Addin:
    
    def __init__(self, xlam_file:str, visible:bool=False) -> None:
        '''The Excel add-in object, including ribbon UI and VBA modules.

        Args:
            xlam_file (str): Add-in file path.
            visible (bool): Process the add-in with Excel application running in the background if False.
        '''
        # work path
        self.xlam_file = xlam_file
        self.path = os.path.dirname(xlam_file)

        # Add-in VBA modules
        self.excel_app = win32com.client.Dispatch('Excel.Application') # win32 COM object
        self.excel_app.Visible = visible
        self.excel_app.DisplayAlerts = False
    

    def close(self):
        '''Close add-in and exit Excel.'''
        self.excel_app.Application.Quit()
    

    def create(self, vba_only:bool=False):
        '''Create addin file.
            - customize ribbon tab and associated VBA callback according to ui file
            - include VBA modules, e.g., general VBA subroutines for data transferring.

        Args:
            vba_only (bool, optional): Whether simple VBA addin (without Python related modules). 
                Defaults to False.
        '''
        N = 2 if vba_only else 3

        # 1 create addin file
        logging.info('(1/%d) Creating add-in structure...', N)
        ui = UI(self.xlam_file)
        template = os.path.join(RESOURCE_PATH, RESOURCE_ADDIN)
        custom_ui = os.path.join(self.path, CUSTOM_UI)
        ui.create(template, custom_ui)

        if not os.path.exists(self.xlam_file):
            raise AddInException('Create add-in structures failed.')

        # 2 update VBA modules
        vba = VBA(xlam_file=self.xlam_file, excel_app=self.excel_app)

        # 2.1 import ribbon module
        logging.info('(2/%d) Creating menu callback subroutines...', N)
        base_menu = os.path.join(RESOURCE_PATH, RESOURCE_VBA, f'{VBA_MENU}.bas')
        user_menu = os.path.join(RESOURCE_PATH, RESOURCE_VBA, f'{VBA_USER_MENU}.bas')
        vba.import_module(base_menu)
        vba.import_module(user_menu)

        # extra steps for VBA-Python combined addin
        if not vba_only:
            logging.info('(3/%d) Creating Python-VBA interaction modules...', N)

            # 2. import general module
            general_module = os.path.join(RESOURCE_PATH, RESOURCE_VBA, f'{VBA_GENERAL}.bas')
            vba.import_module(general_module)

            # 3. copy main python scripts
            if RESOURCE_PATH.upper()!=self.path.upper():
                python_scripts = os.path.join(RESOURCE_PATH, RESOURCE_PYTHON)
                target_scripts = os.path.join(self.path, RESOURCE_PYTHON)
                shutil.copytree(python_scripts, target_scripts)

                python_main = os.path.join(RESOURCE_PATH, PYTHON_MAIN)
                python_config = os.path.join(RESOURCE_PATH, PYTHON_CONFIG)
                shutil.copy(python_main, self.path)
                shutil.copy(python_config, self.path)
        
        # save vba modules
        vba.save()


    def update(self):
        '''Update Ribbon UI. Note that the menu callback functions should be updated manually.'''
        # update addin with customized ui file
        logging.info('(1/1) Updating ribbon structures...')
        custom_ui = os.path.join(self.path, CUSTOM_UI)
        ui = UI(self.xlam_file)
        ui.update(custom_ui)