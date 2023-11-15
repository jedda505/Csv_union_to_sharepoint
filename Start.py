'''
This script will automatically activate the venv for this project without the need for console commands
'''

import subprocess

p1 = subprocess.run(r'".venv\Scripts\activate.bat" & RunGUI.py', shell=True, text=True)

#print(p1.stderr, p1.stdout)
