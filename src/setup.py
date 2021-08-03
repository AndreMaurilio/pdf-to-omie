# Cria executavel compativel com Windows

import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["os"],
    # "zip_include_packages": "PyQt5",
    "includes": ["PyQt5.QtWidgets"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"
setup(
    name="OmieConvert",
    version="0.1",
    description="convertion",
    options={"build_exe": build_exe_options},
    executables=[Executable("window_pdf_xlsx.py", base=base)]
)

# from distutils.core import setup
# import py2exe
#
# includes = ["sip",
#             "PyQt5",
#             "PyQt5.QtCore",
#             "PyQt5.QtGui",
#             "pdfminer.pdfinterp",
#             "pdfminer.layout",
#             "pdfminer.converter",
#             "pdfminer.pdfpage",
#             "io",
#             "mf34.py",
#             "dist.py",
#             "distutils.core",
#             "distutils.dist"
#             ]
#
# setup(windows=['window_pdf_xlsx.py'],
#       options={
#           "py2exe": {
#               "optimize": 2,
#               "includes": includes
#           }
#       }
#       )
