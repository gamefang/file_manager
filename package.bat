@echo off
pyinstaller -F -i icon.ico Main.py
cd dist
del FileManager.exe
ren Main.exe FileManager.exe
cd ..
copy FileManager.xlsx dist\FileManager.xlsx
copy FileManager.json dist\FileManager.json
pause