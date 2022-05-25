import os
from ..share import AddInException


# menu callback function
CB_TEMPLATE = """
Sub {callback}(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
End Sub\n
"""


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


    def add_callbacks(self, module_name:str, list_cb:list, base_module_file:str=''):
        '''Add callback functions to specified module, combined with initial codes saved in
        specified file.

        Args:
            module_name (str): module name to update.
            list_cb (list): callback functions to add.
            base_module_file (str, optional): initial module file to merge.
        '''
        # initial code
        code = ''
        if base_module_file and os.path.isfile(base_module_file):
            with open(base_module_file, 'r') as f:
                code = f.read()
        
        # add callback functions not included in the initial code
        cbs = [CB_TEMPLATE.format(callback=cb) for cb in list_cb if f'Sub {cb}(' not in code]
        code += '\n'.join(cbs)

        # combined vba file
        module_file = os.path.join(os.path.dirname(self.xlam_file), f'{module_name}.bas')
        with open(module_file, 'w') as f:
            f.write(code)
        
        # import module with file
        self.import_module(module_file)
        os.remove(module_file)


    def update_callbacks(self, module_name:str, list_cb:str):
        '''Update module with a list of callback functions. It will be skipped if a function has already 
        existed in the module.

        Args:
            module_name (str): module name to update.
            list_cb (str): callback functions to add.
        '''
        # export current module as a base line
        module_file = os.path.join(os.path.dirname(self.xlam_file), f'{module_name}.bas')
        self.export_module(module_name, module_file)

        # merge code with the input callbacks
        self.add_callbacks(module_name, list_cb, module_file)


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

    