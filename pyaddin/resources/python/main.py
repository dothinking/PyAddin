'''
This is the interface between VBA and python, which is called by VBA function:
```
Dim res as Object
Set res = RunPython("package.module.method", arg1, arg2, ...)
```
Actually the command below is executed in the background:
```
python main.py workbook_name package.module.method arg1 arg2 ...
```
VBA calls python by console arguments, and gets return from temp files, which 
are results of running python.

2019.01
'''

import sys
import os
from scripts.utils import context

MAIN_PATH = os.path.dirname(os.path.abspath(__file__)) 
sys.path.append(MAIN_PATH)


def redirect(fun):
    '''Decorator for user defined function called by VBA.'''    
    def wrapper(*args, **kwargs):
        res = ''
        try:
            res = fun(*args, **kwargs)
        except Exception as e:
            context.get_logger().error(e)
        else:
            context.get_logger().info(res)
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
    module_file = os.path.join(MAIN_PATH, f'{"/".join(modules_name)}.py')

    # import module dynamically if exists
    module_path = '.'.join(modules_name)
    if not os.path.exists(module_file):
        context.get_logger().error('Python module "%s" does not exist.', module_path)
    module = __import__(module_path, fromlist=('ooh'), globals={'name': 100})

    # import method if exists
    if not hasattr(module, method_name):
        context.get_logger().error('Python method "%s" does not exist.', method_name)
    fun = getattr(module, method_name)

    # store caller workbook
    context.set_caller(caller_name)

    return fun(*args)


if __name__ == '__main__':

    # start context
    context.start()

    # redirect output/error to log files
    context.set_logger(MAIN_PATH)

    # python main.py workbook_name package.module.method *args
    run_python_method(sys.argv[1], sys.argv[2], *sys.argv[3:])