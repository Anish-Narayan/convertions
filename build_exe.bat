@echo off
REM Run from this folder on Windows with Python 3.10+ installed.
py -m venv .venv && call .venv\Scripts\activate
pip install -r requirements-converter.txt pyinstaller
pyinstaller --noconfirm --onefile --windowed --name SQLConverter --collect-all sqlglot converter_gui.py
echo Built: dist\SQLConverter.exe
