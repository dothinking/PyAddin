import os
import shutil
import yaml
import win32com.client

SCRIPT_PATH = os.path.dirname(os.path.abspath(__file__)) 
RES_PATH = os.path.join(SCRIPT_PATH, 'res')
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
	'''

	# check
	ui_file = os.path.join(path, CUSTOMUI)
	if not os.path.exists(ui_file):
		raise Exception('Can not find {0} under current path.'.format(CUSTOMUI))

	with open(ui_file, 'r') as f:
		dict_ui = yaml.load(f)

	if not dict_ui:
		raise Exception('Empty {0}'.format(CUSTOMUI))

	# join ui xml
	ui_xml = _create_custom_ui_xml(dict_ui)

	# combine ui xml to addin xml
	source_addin = os.path.join(RES_PATH, RES_ADDIN)
	dest_addin = os.path.join(path, RES_ADDIN)
	ui_path = os.path.join(dest_addin, 'customUI')
	if os.path.isdir(dest_addin):
		shutil.rmtree(dest_addin)
	shutil.copytree(source_addin, dest_addin)

	os.mkdir(ui_path)
	with open(os.path.join(ui_path, 'CustomUI.xml'), 'w') as f:
		f.write(ui_xml)

	# achive and remove original package
	base_name = addin_name.replace('.xlam', '')
	shutil.make_archive(base_name, 'zip', dest_addin)
	shutil.rmtree(dest_addin)	

	# convert to addin *.xlam
	addin_file = '{0}.xlam'.format(base_name)
	shutil.move('{0}.zip'.format(base_name), addin_file)

	# create callback function module for customed menu button
	callbacks = []
	for tab, groups in dict_ui.items():
		for group, btns in groups.items():
			for btn, attrs in btns.items():
				callbacks.append(attrs.get('onAction', None))


	cb_template = """
Sub {callback}(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''     
End Sub\n
"""
	
	menu_module = '{0}.bas'.format(VBA_MENU)
	with open(menu_module, 'w') as f:
		# header
		with open(os.path.join(RES_PATH, RES_VBA, menu_module), 'r') as ff:
			f.write(ff.read())

		# callback functions
		for cb in callbacks:
			f.write(cb_template.format(callback=cb))



	# import menu module to addin
	xl = win32com.client.Dispatch("Excel.Application")
	xl.Visible = 0
	xl.DisplayAlerts = 0

	for x in xl.Workbooks:
		print(x)
	wb = xl.Workbooks.open(os.path.join(path, addin_file))
	wb.VBProject.VBComponents.Import(os.path.join(path, menu_module))

	if not vba_only:
		# thisworkbook
		workbook_module = os.path.join(RES_PATH, RES_VBA, 'ThisWorkbook.cls')
		with open(workbook_module, 'r') as f:
			wb_module_string = f.read()
		xlmodule = wb.VBProject.VBComponents("ThisWorkbook")
		xlmodule.CodeModule.AddFromString(wb_module_string)

		# general module
		general_module = os.path.join(RES_PATH, RES_VBA, '{0}.bas'.format(VBA_GRNERAL))
		wb.VBProject.VBComponents.Import(os.path.join(path, general_module))

	# xl.Application.Run('Module1.test')

	# for c in wb.VBProject.VBComponents:
	# 	if c.Type != 1: continue
	# 	print(c.Name)
	# 	num = c.CodeModule.CountOfLines
	# 	src = str(c.CodeModule.Lines(0,num)).split('\n')
	# 	for line in src:
	# 		print(line)
	# 	print('*'*20, '\n')

	xl.DisplayAlerts = False
	wb.DoNotPromptForConvert = True
	wb.CheckCompatibility = False

	wb.Save()
	xl.Application.Quit()

	os.remove(menu_module)

	if vba_only:
		return



	

	# print(ui_xml)






def update_addin():
	pass




def _create_custom_ui_xml(dict_ui):
	'''join CustomUI.xml'''
	ui_xml = '''
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon startFromScratch="false">
    <tabs>
    {0}
    </tabs>
  </ribbon>
</customUI>'''
	
	# tab
	tabs_xml = ''
	for i, (tab, groups) in enumerate(dict_ui.items(), start=1):
		# group
		groups_xml = ''
		for j, (group, btns) in enumerate(groups.items(), start=1):
			# button
			btns_xml = ''
			for k, (btn, attrs) in enumerate(btns.items(), start=1):
				btn_attrs = ['{0}="{1}"'.format(key,val) for key,val in attrs.items() if val]
				btns_xml += '<button id="btn_{0}_{1}_{2}" label="{3}" {4}/>\n'.format(i, j, k, btn, ' '.join(btn_attrs))
			groups_xml += '<group id="group_{0}_{1}" label="{2}">\n{3}</group>\n'.format(i, j, group, btns_xml)
		tabs_xml += '<tab id="tab_{0}" label="{1}">\n{2}</tab>\n'.format(i, tab, groups_xml)	

	return ui_xml.format(tabs_xml)