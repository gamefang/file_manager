@echo off
pyinstaller -F -w filebat.py
copy filebat.xlsx dist\filebat.xlsx
pause