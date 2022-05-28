import shutil
import os


class AddInException(Exception):
    '''User defined Exception.'''
    pass



def copytree(from_dir:str, to_dir:str):
    '''Copy all files and folders to specified folder. Note the destination folder
    is existed, while ``shutil.copytree`` works when destination folder doesn't exist.

    Args:
        from_dir (str): from folder.
        to_dir (str): to folder.
    '''
    for filename in os.listdir(from_dir):
        item = os.path.join(from_dir, filename)
        # copy file
        if os.path.isfile(item):
            shutil.copy(item, to_dir)
        # copy folder
        else:
            to_item = os.path.join(to_dir, filename)
            if os.path.isdir(to_item): 
                shutil.rmtree(to_item)
            shutil.copytree(item, to_item)