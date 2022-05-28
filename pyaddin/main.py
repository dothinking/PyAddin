import os
import shutil
import logging
from .addin import (Addin, RESOURCE_PATH, CUSTOM_UI)


class PyAddin:
    '''Command line interface for ``PyAddin``.'''

    @staticmethod
    def init(name:str, vba:bool=False, quiet:bool=True):
        '''Create template project with specified name under current path.
        
        Args:
            name (str) : the name of add-in to create (without the suffix ``.xlam``).
            vba (bool): create VBA add-in only if True, otherwise VBA-Python addin by default.
            quiet (bool): perform the process in the background if True.
        '''
        # new project
        work_path = os.getcwd()
        project_path = os.path.join(work_path, name)
        if os.path.exists(project_path): 
            logging.error(f'Project {name} already existed.')
            return
        os.mkdir(project_path)

        # template UI file
        ui_file = os.path.join(RESOURCE_PATH, CUSTOM_UI)        
        shutil.copy(ui_file, project_path)

        # template add-in based on UI template
        filename = os.path.join(project_path, f'{name}.xlam')
        addin = Addin(xlam_file=filename, visible=not quiet)

        try:
            addin.create(vba_only=vba)
        except Exception as e:
            logging.error(e)
            addin.close()
        else:
            if quiet: addin.close()


    @staticmethod
    def update(name:str, quiet:bool=True):
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