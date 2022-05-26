# this is the interface between VBA and python, which is called
# by VBA function:
# 
# args = Array(arg1, arg2, ...)
# true_or_false = RunPython("package.module.method", args, res)
# 
# actually the command below is executed in the background:
# 
# python main.py "package.module.method" arg1 arg2 ...
# 
# VBA calls python by console arguments, and gets return from 
# temp files, which are results of running python
# 
# 2019.01
# 

import sys
import os
import win32com.client
import context

main_path = os.path.dirname(os.path.abspath(__file__)) 
sys.path.append(main_path)


def redirect(fun):
    '''decorator for user defined function called by VBA'''
    
    def wrapper(*args, **kwargs):
        res = None
        try:
            res = fun(*args, **kwargs)
        except Exception as e:
            sys.stderr.write(str(e))
        else:
            if res: sys.stdout.write(str(res))
        return res

    return wrapper


@redirect
def run_python_method(caller_name:str, key:str, *args):
    '''call method specified by key with arguments: args

    Args:
        caller_name (str): Workbook name calling this script.
        key (str): script path: package.module.method.
    '''
    *modules_name, method_name = key.split('.')
    module_file = os.path.join(main_path, f'{"/".join(modules_name)}.py')

    # import module dynamically if exists
    module_path = '.'.join(modules_name)
    assert os.path.exists(module_file), 'Error Python module "{0}"'.format(module_path)
    module = __import__(module_path, fromlist=('blabla'), globals={'name': 100})

    # import method if exists
    assert hasattr(module, method_name), 'Error Python method "{0}"'.format(method_name)
    fun = getattr(module, method_name)

    # store caller workbook
    wb = get_caller_workbook(caller_name)
    context.set_caller(wb)

    return fun(*args)


class Logger(object):
    '''redirect standard output/error to files, which are bridges for
    communication between python and VBA
    '''
     
    def __init__(self, log_file="out.log", terminal=sys.stdout):
        self.terminal = terminal
        self.log = open(log_file, "w")
 
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
 
    def flush(self):
        pass


def get_caller_workbook(name:str):
    '''Get Workbook instance (win32com) by name.

    Args:
        name (str): Workbook name.
    '''
    app = win32com.client.Dispatch('Excel.Application')
    for wb in app.Workbooks:
        if wb.Name == name: return wb
    return None




if __name__ == '__main__':

    # start context
    context.start()

    # get output folder name from main.cfg
    output = 'outputs'
    output_file, error_file = "output.log", "errors.log"
    with open(os.path.join(main_path, 'main.cfg'), 'r') as f:
        while True:
            line = f.readline().strip()
            if not line:
                break
            elif line.startswith('[output]'):
                output = f.readline().strip()
            elif line.startswith('[stdout]'):
                output_file = f.readline().strip()
            elif line.startswith('[stderr]'):
                error_file = f.readline().strip()

    output_path = os.path.join(main_path, output)
    if not os.path.exists(output_path):
        os.mkdir(output_path)

    # redirect output/error to output.log/errors.log
    sys.stdout = Logger(os.path.join(output_path, output_file), sys.stdout)
    sys.stderr = Logger(os.path.join(output_path, error_file), sys.stderr)

    # python main.py package.module.method *args
    run_python_method(sys.argv[1], sys.argv[2], *sys.argv[3:])