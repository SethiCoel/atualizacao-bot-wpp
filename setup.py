import sys
from cx_Freeze import setup, Executable


additional_files = ['update.py']


#Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], "includes": [], "include_files":[]}


#GUI applications require a different base on Windows (the default is for
#a console application).
# base = None
# if sys.platform == "win32":
#     base = "Win32GUI"



setup(
    name="AutomaBot.Ink",
    version="1.6",
    description="Faz automação de mensagens no WhatsApp",
    options={"build_exe": build_exe_options},
    executables=[Executable("Mensagem.Automatica.py", base=None, icon="icon.ico"),
                 
                 ]
)



# ["aspose-cells",
# "attrs",
# "certifi",
# "cffi",
# "charset-normalizer",
# "cx-Logging",
# "cx_Freeze",
# "et-xmlfile",
# "h11",
# "idna",
# "JPype1",
# "lief",
# "openpyxl",
# "outcome",
# "packaging",
# "pycparser",
# "PySocks",
# "python-dotenv",
# "requests",
# "selenium",
# "sniffio",
# "sortedcontainers",
# "trio",
# "trio-websocket",
# "typing_extensions",
# "urllib3",
# "webdriver-manager"
# "wsproto"]