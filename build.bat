@echo off
pyinstaller --onefile --noconsole ^
--icon=assets/SEV.ico ^
--add-data "assets/*;assets" ^
--add-data "parsers/*;parsers" ^
--name NewsParser.exe main.py
pause