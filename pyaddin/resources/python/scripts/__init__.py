import sys
import os

# solve the issue:
# ValueError: attempted relative import beyond top-level package
script_path = os.path.dirname(os.path.abspath(__file__)) 
sys.path.append(os.path.dirname(script_path))