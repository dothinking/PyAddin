import os
import shutil
import xml.etree.ElementTree as ET
import win32com.client
from .pyvba import UICreator, VBAWriter

SCRIPT_PATH = os.path.dirname(os.path.abspath(__file__)) 
RES_PATH = os.path.join(os.path.dirname(SCRIPT_PATH), 'res')
RES_ADDIN = 'addin'
RES_PYTHON = 'python'
RES_VBA = 'vba'
CUSTOMUI = 'CustomUI.xml'
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

    # check CustomUI.xml -> get callback functions
    callbacks = _get_callbacks_from_CustomUI(path)

    # create addin with customed ui
    addin = UICreator(path, addin_name)
    addin.create(os.path.join(RES_PATH, RES_ADDIN), os.path.join(path, CUSTOMUI))

    if not os.path.exists(addin.addin_file):
        raise Exception('Create Addin structures failed.')

    # VBA writer
    vba = VBAWriter(addin.addin_file)
    try:
        # import menu module
        # create callback function module for customed menu button
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
    callbacks = _get_callbacks_from_CustomUI(path)

    # create addin with customed ui
    addin = UICreator(path, addin_name)
    addin.update(os.path.join(path, CUSTOMUI))

    if not os.path.exists(addin.addin_file):
        raise Exception('Update Addin custom UI failed.')

    # VBA writer
    vba = VBAWriter(addin.addin_file)
    try:
        # update menu module
        # get new callback functions
        vba.update_callbacks(VBA_MENU, callbacks)

    except Exception as e:
        raise e
    finally:
        vba.quit()


def _get_callbacks_from_CustomUI(path):
    '''parse CustomUI.xml to collect all callback function names -> attribute=onAction'''

    ui_file = os.path.join(path, CUSTOMUI)
    if not os.path.exists(ui_file):
        raise Exception('Can not find {0} under current path.'.format(CUSTOMUI))
    else:
        try:
            tree = ET.parse(ui_file)
        except ET.ParseError as e:
            raise Exception('Error format in {0}: {1}'.format(CUSTOMUI, str(e)))
        else:
            root = tree.getroot()

    # get root and check all nodes by iteration
    attr_name = 'onAction'    
    callbacks = [node.attrib.get(attr_name) for node in root.iter() if attr_name in node.attrib]

    if not callbacks:
        raise Exception('Please check {0}: no defined actions'.format(CUSTOMUI))

    return callbacks

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