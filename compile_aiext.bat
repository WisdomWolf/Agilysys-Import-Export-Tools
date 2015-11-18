@Echo OFF

pyinstaller --onefile "Agilysys Import Export Tool.spec"
CD dist
7z a AIEXT.zip *.exe
PAUSE
DEL *.exe