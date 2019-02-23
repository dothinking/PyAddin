import os
import shutil
import win32com.client

class UICreator:
    ''' create Excel addin with customed ribbon tab via python '''

    def __init__(self, path, name='addin'):
        '''
        :param path: addin directory
        :param name: addin name -> {name}.xlam
        :param py_mode: work on both Python and VBA, otherwise VBA mode only
        '''

        self.path = path
        self.name = name.replace('.xlam', '')
        self.addin_file = os.path.join(self.path, '{0}.xlam'.format(self.name))

    def create(self, xml_path, xml_customui):
        '''combine CustomUI.xml with source xml files
        :param xml_path: directory for template xml files extracted from base addin *.xlam
        :param xml_customui: CustomUI.xml
        '''
        # e.g. xml_path = /path/to/xnl_dir ->
        # dest_xml_path = self.path/xml_dir
        dest_xml_path = os.path.join(self.path, os.path.basename(xml_path))

        # if source xml files are not in current path, copy them to current path
        # otherwise deal with source xml files directly
        if dest_xml_path != xml_path:
            # delete dest dir if exists, otherwise the copy will be forbidden
            if os.path.isdir(dest_xml_path):
                shutil.rmtree(dest_xml_path)
            shutil.copytree(xml_path, dest_xml_path)

        # customUI under destination path
        dest_ui_path = os.path.join(dest_xml_path, 'customUI')
        if not os.path.exists(dest_ui_path):
            os.mkdir(dest_ui_path)
        shutil.copy(xml_customui, dest_ui_path)

        # achive xml files and remove original package
        shutil.make_archive(os.path.join(self.path, self.name), 'zip', dest_xml_path)
        shutil.rmtree(dest_xml_path)

        # convert to addin *.xlam
        zip_file = '{0}.zip'.format(self.name)
        shutil.move(os.path.join(self.path, zip_file), self.addin_file)

    def update(self, xml_customui):
        '''update current addin ribbon with CustomUI.xml defined by xml_customui'''

        if not os.path.exists(self.addin_file):
            raise Exception('Current addin does not exist yet.')

        # unpack
        zip_file = os.path.join(self.path, '{0}.zip'.format(self.name))
        xml_path = os.path.join(self.path, self.name)
        shutil.move(self.addin_file, zip_file)
        shutil.unpack_archive(zip_file, xml_path)

        # recreate
        self.create(xml_path, xml_customui)

class VBAWriter:

    def __init__(self, addin_file='addin.xlam'):
        '''import vba models for addin file
        :param addin_file: addin directory
        '''

        assert os.path.exists(addin_file), "Can not find Excel addin: {0}".format(addin_file)
        self.addin_file = addin_file

        # win32 COM object
        self.xl = win32com.client.Dispatch("Excel.Application")
        self.xl.Visible = 0
        self.xl.DisplayAlerts = 0
        self.xl.DisplayAlerts = False

        # current workbook
        self.wb = self.xl.Workbooks.open(addin_file)
        self.wb.DoNotPromptForConvert = True
        self.wb.CheckCompatibility = False

    def add_callbacks(self, module_name, list_cb, combined_module_file=''):
        '''create module with initial codes combined with a list of callback functions
        :param module_name: module name        
        :param list_cb: callback functions to be imported
        :param combined_module_file: initial module file to be merged
        ''' 
        # initial codes from combined_module_file
        if combined_module_file and os.path.isfile(combined_module_file):
            with open(combined_module_file, 'r') as f:
                initial_codes = f.read()
        else:
            initial_codes = ''

        # callback functions
        cb_template = """
Sub {callback}(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''     
End Sub\n
""" 
        module_file = os.path.join(os.path.dirname(self.addin_file), '{0}.bas'.format(module_name))
        with open(module_file, 'w') as f:
            # initial codes
            if initial_codes:
                f.write(initial_codes)

            # callback functions
            for cb in list_cb:
                if not (initial_codes and 'Sub {0}('.format(cb) in initial_codes):
                    f.write(cb_template.format(callback=cb))

        self.import_module(module_file)
        os.remove(module_file)


    def update_callbacks(self, module_name, list_cb):
        '''update module with a list of callback functions.
        if a function has already existed in the module, it will be skipped'''

        # export current module as a base line
        module_file = os.path.join(os.path.dirname(self.addin_file), '{0}.bas'.format(module_name))
        self.export_module(module_name, module_file)

        # merge the input callbacks
        self.add_callbacks(module_name, list_cb, module_file)


    def import_module(self, module_file):
        '''import vba module from *.bas. module name is declared in the first line of module_file:
            Attribute VB_Name = "xxxx"
        '''
        # check module name
        with open(module_file, 'r') as f:
            header = f.readline()
        if header.startswith('Attribute VB_Name'):
            module_name = header.split('"')[-2]
        else:
            raise Exception('The first line of a valid module file should be: Attribute VB_Name = "xxxx"')

        # check modules
        for comp in self.wb.VBProject.VBComponents:
            if comp.Name == module_name:
                self.wb.VBProject.VBComponents.Remove(comp)
                break
            
        self.wb.VBProject.VBComponents.Import(module_file)

    def import_named_module(self, module_name, module_file):
        '''import vba code for named module
        :param module_name: module name, e.g. sheet1, sheet2, thisworkbook
        :param module_file: module file
        '''
        code_module = self.wb.VBProject.VBComponents(module_name).codeModule

        # clear contents if already existing
        num = code_module.CountOfLines
        if num:
            code_module.DeleteLines(1, num)
            code_module.AddFromFile(module_file)


    def run_macro(self, macro_name):
        '''run macro embedded in current addin
        :param macro_name: macro name, e.g. mudule.sub_name
        '''
        self.xl.Application.Run(macro_name)

    def export_module(self, module_name, saved_file):
        '''export vba module to file'''
        for comp in self.wb.VBProject.VBComponents:
            if comp.Name == module_name:
                comp.Export(saved_file)
                return True
        else:
            return False

    def read_module(self, module_name):
        '''get source codes from module with name module_name'''

        for comp in self.wb.VBProject.VBComponents:
            if comp.Name == module_name:
                num = comp.CodeModule.CountOfLines
                src = comp.CodeModule.Lines(1,num)
                break
        else:
            src = None

        return src

    def quit(self):
        self.wb.Save()
        self.xl.Application.Quit()


if __name__ == '__main__':

    import traceback
   
    SCRIPT_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    RES_PATH = os.path.join(SCRIPT_PATH, 'res')
    RES_ADDIN = os.path.join(RES_PATH, 'addin')
    RES_PYTHON = os.path.join(RES_PATH, 'python')
    RES_VBA = os.path.join(RES_PATH, 'vba')
    CUSTOMUI = 'CustomUI.xml'
    VBA_GRNERAL = 'general'
    VBA_MENU = 'menu'

    path = os.path.join(os.path.dirname(SCRIPT_PATH), 'test')
    ui_file = os.path.join(path, CUSTOMUI)
    callbacks = ['hello_word', 'about']

    # create addin with customed ui
    addin = UICreator(path, 'myaddin')
    addin.create(RES_ADDIN, ui_file)

    # VBA writer
    vba = VBAWriter(addin.addin_file)
    try:
        # import menu module
        vba.add_callbacks(VBA_MENU, callbacks, os.path.join(RES_VBA, '{0}.bas'.format(VBA_MENU)))

        # import workbook module
        workbook_module = os.path.join(RES_VBA, 'ThisWorkbook.cls')
        vba.import_named_module("ThisWorkbook", workbook_module)

        # import general module
        general_module = os.path.join(RES_VBA, '{0}.bas'.format(VBA_GRNERAL))
        vba.import_module(general_module)

        # update module
        callbacks.append('test02')
        vba.update_callbacks(VBA_MENU, callbacks)

    except Exception as e:
        traceback.print_exc()
    finally:
        vba.quit()


