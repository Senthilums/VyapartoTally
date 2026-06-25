@echo off
call .venv\Scripts\activate
pip install -r requirements.txt

rmdir /s /q build
rmdir /s /q dist
del VyaparToTally.spec

pyinstaller --onefile --windowed --name VyaparToTally ui.py

pause