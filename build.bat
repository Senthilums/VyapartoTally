@echo off
call .venv\Scripts\activate
pip install -r requirements_vyapar_daybook_to_tally.txt

rmdir /s /q build
rmdir /s /q dist

pyinstaller VyaparToTally.spec

pause
