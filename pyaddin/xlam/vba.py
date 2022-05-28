import os
from ..share import AddInException


class VBA:

    def __init__(self, xlam_file:str, excel_app):
        '''Process VBA modules for add-in file, e.g., add callback function based on ribbon menu.

        Args:
            xlam_file (str): Add-in filename.
            excel_app (win32 COM object): The Excel application instance.
        '''
        self.xlam_file = xlam_file
        if not os.path.exists(xlam_file):
            raise AddInException(f'Not find add-in file: {self.xlam_file}.')

        # current workbook
        self.wb = excel_app.Workbooks.open(xlam_file)
        self.wb.DoNotPromptForConvert = True
        self.wb.CheckCompatibility = False


    def save(self):
        '''Save add-in.'''
        self.wb.Save()


    def import_module(self, module_file:str):
        '''Import vba module from specified file (*.bas). 

        Args:
            module_file (str): Module file with module name declared in the first line: 
                Attribute VB_Name = "xxxx"
        '''
        # check module name
        with open(module_file, 'r') as f:
            header = f.readline().strip()

        if not header.startswith('Attribute VB_Name'):
            raise AddInException('The first line of a valid module file should be: Attribute VB_Name = "xxxx"')

        # delete module if existed already
        module_name = header.split('"')[-2]
        for comp in self.wb.VBProject.VBComponents:
            if comp.Name == module_name:
                self.wb.VBProject.VBComponents.Remove(comp)
                break

        # import module
        self.wb.VBProject.VBComponents.Import(module_file)


    def import_named_module(self, module_name:str, module_file:str):
        '''Import vba code for specified module.

        Args:
            module_name (str): module name, e.g., sheet1, sheet2, thisWorkbook
            module_file (str): Module file with module name declared in the first line: 
                Attribute VB_Name = "xxxx"
        '''
        code_module = self.wb.VBProject.VBComponents(module_name).codeModule

        # clear contents if already existing
        num = code_module.CountOfLines
        if num:
            code_module.DeleteLines(1, num)
            code_module.AddFromFile(module_file)


    def export_module(self, module_name:str, export_filename:str) -> bool:
        '''Export VBA module to specified file.

        Args:
            module_name (str): module name to export.
            export_filename (str): filename for the exported codes.

        Returns:
            bool: exporting status, success or failed.
        '''
        for comp in self.wb.VBProject.VBComponents:
            if comp.Name == module_name:
                comp.Export(export_filename)
                return True
        return False


    def read_module(self, module_name:str) -> str:
        '''Get source codes from specified module.

        Args:
            module_name (str): VBA module name to extract codes.

        Returns:
            str: exported codes.
        '''
        src = None
        for comp in self.wb.VBProject.VBComponents:
            if comp.Name == module_name:
                num = comp.CodeModule.CountOfLines
                src = comp.CodeModule.Lines(1,num)
                break

        return src

    