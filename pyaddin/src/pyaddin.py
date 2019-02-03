import os
import shutil
import yaml
import win32com.client
from .pyvba import UICreator, VBAWriter

SCRIPT_PATH = os.path.dirname(os.path.abspath(__file__)) 
RES_PATH = os.path.join(os.path.dirname(SCRIPT_PATH), 'res')
RES_ADDIN = 'addin'
RES_PYTHON = 'python'
RES_VBA = 'vba'
CUSTOMUI = 'customUI.yaml'
VBA_GRNERAL = 'general'
VBA_MENU = 'menu'


def init_project(path):
    '''initialize ui config file under path
    '''

    ui_file = os.path.join(RES_PATH, CUSTOMUI)
    shutil.copy(ui_file, path)

def create_addin(path, addin_name='addin', vba_only=False):
    '''create addin:
        - customize ribbon tab and associated VBA callback according to ui file
        - include VBA modules for VBA-Python addin
        :param path: path for the addin to be created
        :param addin_name: name of the addin to be created
    '''

    # parse UI dict from customed file
    dict_ui = _parse_UI(path)

    # create addin with customed ui
    addin = UICreator(path, addin_name)
    addin.create(os.path.join(RES_PATH, RES_ADDIN), dict_ui)

    if not os.path.exists(addin.addin_file):
        raise Exception('Create Addin structures failed.')

    # VBA writer
    vba = VBAWriter(addin.addin_file)
    try:
        # import menu module
        # create callback function module for customed menu button
        callbacks = []
        for tab, groups in dict_ui.items():
            for group, btns in groups.items():
                for btn, attrs in btns.items():
                    callbacks.append(attrs.get('onAction', None))
        vba.add_callbacks(VBA_MENU, callbacks, os.path.join(RES_PATH, RES_VBA, '{0}.bas'.format(VBA_MENU)))

        # extra steps for VBA-Python combined addin
        if not vba_only:
            # import workbook module
            workbook_module = os.path.join(RES_PATH, RES_VBA, 'ThisWorkbook.cls')
            vba.import_named_module("ThisWorkbook", workbook_module)

            # import general module
            general_module = os.path.join(RES_PATH, RES_VBA, '{0}.bas'.format(VBA_GRNERAL))
            vba.import_module(general_module)

            # copy main python scripts
            _copy_all(os.path.join(RES_PATH, RES_PYTHON), path)

    except Exception as e:
        raise e
    finally:
        vba.quit()

def update_addin(path, addin_name='addin'):
    '''update Ribbon Tab and associated callback functions for addin with `addin_name` 
    under `path` according to `customUI.yaml`
    ：param path: addin path
    :param addin_name: name of addin to be updated under current `path`
    '''

    # parse UI dict from customed file
    dict_ui = _parse_UI(path)

    # create addin with customed ui
    addin = UICreator(path, addin_name)
    addin.update(dict_ui)

    if not os.path.exists(addin.addin_file):
        raise Exception('Update Addin ribbon tab structures failed.')

    # VBA writer
    vba = VBAWriter(addin.addin_file)
    try:
        # update menu module
        # get new callback functions
        callbacks = []
        for tab, groups in dict_ui.items():
            for group, btns in groups.items():
                for btn, attrs in btns.items():
                    callbacks.append(attrs.get('onAction', None))

        vba.update_callbacks(VBA_MENU, callbacks)

    except Exception as e:
        raise e
    finally:
        vba.quit()


def _parse_UI(path):
    '''parse UI dict from yaml file'''
    ui_file = os.path.join(path, CUSTOMUI)
    if not os.path.exists(ui_file):
        raise Exception('Can not find {0} under current path.'.format(CUSTOMUI))

    with open(ui_file, 'r') as f:
        try:
            dict_ui = yaml.load(f)
        except yaml.YAMLError as e:
            raise Exception('Error format for UI configuration file: {0}'.format(str(e))) 

    if not dict_ui:
        raise Exception('Empty {0}'.format(CUSTOMUI))

    return dict_ui

def _copy_all(path, out):
    '''copy all files and dirs under path to out
    '''
    for files in os.listdir(path):
        name = os.path.join(path, files)
        back_name = os.path.join(out, files)
        if os.path.isfile(name):
            shutil.copy(name, back_name)
        else:
            if not os.path.isdir(back_name):
                os.makedirs(back_name)
            _copy_all(name, back_name)