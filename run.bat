@echo off
py -c "import requests, openpyxl" 2>nul
if errorlevel 1 (
    py -m pip install -r requirements.txt
)
py main.py
pause
