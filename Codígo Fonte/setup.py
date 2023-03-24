import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["os", "PySimpleGUI", "translate", "openpyxl"],
    "include_files": ["icon/"]
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Tradutor de Planilhas Excel",
    version="0.1",
    description="Um programa para traduzir planilhas Excel",
    options={"build_exe": build_exe_options},
    executables=[Executable("MenoRobot.py", base=base, icon="icon/icon.ico")]
)
