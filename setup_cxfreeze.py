#setup.py
import sys, os
from cx_Freeze import setup, Executable

__version__ = "1.1.0"

#excludes = ["tkinter", "PyQt4.QtSql", "sqlite3", "scipy.lib.lapack.flapack",
#            "PyQt4.QtNetwork", "PyQt4.QtScript", "PyQt5"]
excludes = ["tkinter"]

packages = ["numpy"]


setup(
    name = "Caliber",
    description='The System Caliber. Price transformations for Goodwin',
    version=__version__,
    options = {"build_exe": {
            'packages': packages,
            'excludes': excludes,
            'include_msvcr': True,
             "optimize": 2
            }},
    executables = [Executable("Caliber.py",base="Win32GUI")]
)
