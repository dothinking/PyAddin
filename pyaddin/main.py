import os
import shutil
import logging
from .addin import (Addin, RESOURCE_PATH, CUSTOM_UI)


class PyAddin:
    '''Command line interface for ``PyAddin``.'''

    @staticmethod
    def init():
        '''Initialize project and set current path as working path.'''
        ui_file = os.path.join(RESOURCE_PATH, CUSTOM_UI)
        work_path = os.getcwd()
        shutil.copy(ui_file, work_path)
    

    @staticmethod
    def create(name:str='addin', vba:bool=False, quiet:bool=True):
        '''Create add-in file (name.xlam) based on ribbon UI file (CustomUI.xml) under working path.
        
        Args:
            name (str) : the name of add-in to create (without the suffix ``.xlam``).
            vba (bool): create VBA add-in only if True, otherwise VBA-Python addin by default.
            quiet (bool): perform the process in the background if True.
        '''
        filename = os.path.join(os.getcwd(), f'{name}.xlam')
        addin = Addin(xlam_file=filename, visible=not quiet)

        try:
            addin.create(vba_only=vba)
        except Exception as e:
            logging.error(e)
            addin.close()
        else:
            if quiet: addin.close()


    @staticmethod
    def update(name:str='addin', quiet:bool=True):
        '''Update add-in file (name.xlam) based on ribbon UI file (CustomUI.xml) under working path.
        
        Args:
            name (str) : the name of add-in to update (without the suffix ``.xlam``).
            quiet (bool): perform the process in the background if True.
        '''
        filename = os.path.join(os.getcwd(), f'{name}.xlam')
        addin = Addin(xlam_file=filename, visible=not quiet)

        try:
            addin.update()
        except Exception as e:
            logging.error(e)
            addin.close()
        else:
            if quiet: addin.close()


def main():
    import fire
    fire.Fire(PyAddin)


if __name__ == '__main__':
    main()