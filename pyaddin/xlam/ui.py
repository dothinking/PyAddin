import os
import shutil
from ..share import AddInException


class UI:

    def __init__(self, xlam_file:str) -> None:
        '''Create Excel add-in (*.xlam) with customized ribbon tab.

        Args:
            xlam_file (str): The add-in file to create or update.
        '''
        self.xlam_file = xlam_file
        self.path, self.name = os.path.split(xlam_file)


    def create(self, template_path:str, custom_ui_filename:str='CustomUI.xml'):
        '''Combine CustomUI.xml with template xml files to create a new add-in file.

        Args:
            template_path (str): directory for template xml files extracted from a general add-in file.
            custom_ui_filename (str): full path to CustomUI.xml defining the ribbon UI.
        '''
        # copy template files to current path if they're not under current path
        work_path = os.path.join(self.path, os.path.basename(template_path))
        if work_path.upper() != template_path.upper():
            # delete dest dir if exists, otherwise the copy will be forbidden
            if os.path.isdir(work_path): shutil.rmtree(work_path)
            shutil.copytree(template_path, work_path)

        # copy customUI.xml to path/template/customUI
        dest_ui_path = os.path.join(work_path, 'customUI')
        if not os.path.exists(dest_ui_path):
            os.mkdir(dest_ui_path)
        shutil.copy(custom_ui_filename, dest_ui_path)

        # archive xml files and remove original package
        shutil.make_archive(os.path.join(self.path, self.name), 'zip', work_path)
        shutil.rmtree(work_path)

        # convert to add-in *.xlam
        zip_file = os.path.join(self.path, f'{self.name}.zip')
        shutil.move(zip_file, self.xlam_file)


    def update(self, custom_ui_filename:str='CustomUI.xml'):
        '''Update current add-in ribbon with specified CustomUI.xml.

        Args:
            custom_ui_filename (str): full path to CustomUI.xml defining the ribbon UI.
        '''
        if not os.path.exists(self.xlam_file):
            raise AddInException(f'Not find add-in file: {self.xlam_file}.')

        # rename: name.xlam -> name.zip
        zip_file = os.path.join(self.path, '{0}.zip'.format(self.name))
        shutil.move(self.xlam_file, zip_file)

        # unpack: name.zip -> name/
        xml_path = os.path.join(self.path, self.name)        
        shutil.unpack_archive(zip_file, xml_path)

        # recreate: update current template files with new CustomUI.xml
        self.create(xml_path, custom_ui_filename)
