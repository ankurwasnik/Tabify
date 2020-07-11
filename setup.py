from cx_Freeze import setup, Executable
import sys,os
os.environ['TCL_LIBRARY']=r'C:/Programs/Python/Python35/tcl/tcl8.6'
os.environ['TK_LIBRARY']=r'C:/Programs/Python/Python35/tcl/tk8.6'
setup(name="tabify",
      version="1.0",
      description="Extract tables from pdf.",
      executables=[Executable("testfile.py")]
      )