'''Manage parameters context over modules.'''

KEY_CALLER = 'caller'


def start():
    global _dict
    _dict = {}

def set(key, value):
    _dict[key] = value

def get(key, default_value=None):
    if key in _dict:
        return _dict[key]
    else:
        return default_value


def set_caller(caller):
    '''set workbook instance calling this script.'''
    set(KEY_CALLER, caller)


def get_caller():
    '''get workbook instance calling this script.'''
    return get(KEY_CALLER)