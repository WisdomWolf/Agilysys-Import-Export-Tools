@Echo OFF
rem ***** get rid of all the old files in the build folder
RD /S /Q build

rem ***** ensure clean dist folder
RD /S /Q dist

python setup.py py2exe
CD dist
7z a AIEXT.zip *.*
PAUSE
rem ***** cleanup temp files
DEL *.exe
DEL *.dll
DEL library.zip

rem ***** Need to adjust to include folders in zip