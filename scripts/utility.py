# 
# common functions to be reused
# 
# 2019.01
# 

import os
import sys


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


def udf(fun):
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

    # set a tag that fun is decorated
    setattr(wrapper, 'UDF', True)

    return wrapper