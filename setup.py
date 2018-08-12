import os
from cx_Freeze import setup, Executable

os.environ['TCL_LIBRARY'] = r'C:\python\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\python\tcl\tk8.6'
setup(name = "Opener", 
        version = "0.1", 
        description = "This is a test opener file", 
        executables= [Executable("start.py")] 
    )
