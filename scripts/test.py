# test module
from .utility import catch_exception


@catch_exception
def division(a, b):
	assert a!='', 'cell A1 is empty'
	assert b!='', 'cell A2 is empty'
	return float(a)/float(b)