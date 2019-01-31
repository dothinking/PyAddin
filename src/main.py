import os
import sys
import argparse
import yaml

from pyaddin import init_project, create_addin, update_addin


def main():
	'''commands:
		pyaddin init
		pyaddin create --name --vba
		pyaddin update --path
	'''		

	if len(sys.argv) == 1:
	    sys.argv.append('--help')

	# parse arguments
	parser = argparse.ArgumentParser()
	parser.add_argument('operation', choices=['init', 'create', 'update'], help='init, create, update')
	parser.add_argument('-n','--name', default='addin', help='addin file name to be created: [name].xlam')
	parser.add_argument('-v','--vba', action='store_true', help='create VBA addin only, otherwise VBA-Python addin by default')
	parser.add_argument('-p','--path', default='addin.xlam', help='path of addin file to be updated')
	args = parser.parse_args()

	# do what you need
	current_path = os.getcwd()
	if args.operation == 'init':		
		init_project(current_path)
	elif args.operation == 'create':
		create_addin(current_path, args.name, args.vba)
	elif args.operation == 'update':
		update_addin(current_path, args.path)

if __name__ == '__main__':
	try:
		main()
	except yaml.YAMLError as e:
	    print('Error format for UI configuration file: {0}'.format(str(e)))	        
	except Exception as e:
		print(e)