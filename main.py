# this is the interface between VBA and python, which is called
# by VBA function:
# 
# args = Array("package.module.method", arg1, arg2, ...)
# res = RunPython(args)
# 
# actually the command below is executed in the background:
# 
# python main.py "package.module.method" arg1 arg2 ...
# 
# VBA calls python by console arguments, and gets return from 
# temp files, which are results of running python
# 
# 2019.01

import sys
import os

main_path = os.path.dirname(os.path.abspath(__file__)) 
sys.path.append(main_path)

from scripts import utility


if __name__ == '__main__':

	# python main.py package.module.method *args
	key, *args = sys.argv[1:]

	m = key.split('.')
	module_file = os.path.join(main_path, '{0}.py'.format('/'.join(m[:-1])))
	if os.path.exists(module_file):

		# import module dynamically
		module = __import__('.'.join(m[:-1]), fromlist=True)

		# import method
		if hasattr(module, m[-1]):
			f = getattr(module, m[-1])
			res = f(*args)
			if res != None:
				sys.stdout.write(str(res))
		else:
			sys.stderr.write('Error Python method "{0}"'.format(m[-1]))
	else:
		sys.stderr.write('Error Python module "{0}"'.format('.'.join(m[:-1])))