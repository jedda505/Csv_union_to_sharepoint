'''
This script will automatically install packages via pip to a venv
'''
import os
import subprocess
from venv import EnvBuilder

# Create venv, pip installs

environment = EnvBuilder(with_pip = True)

environment.create(".venv")

p1 = subprocess.run('".venv/Scripts/activate.bat" & pip install -r requirements.txt', shell=True, capture_output = True, text = True)

print(p1.returncode, p1.stdout, p1.stderr)
