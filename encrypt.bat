@echo off

REM Check if Flask is installed
pip show Flask >nul 2>&1
if errorlevel 1 (
    REM Flask is not installed, install it
    pip install Flask
)

REM Check if openpyxl is installed
pip show openpyxl >nul 2>&1
if errorlevel 1 (
    REM openpyxl is not installed, install it
    pip install openpyxl
)

REM Check if cryptography is installed
pip show cryptography >nul 2>&1
if errorlevel 1 (
    REM cryptography is not installed, install it
    pip install cryptography
)

REM Navigate to the src folder
cd src

REM Run the Python script
python encrypt.py

pause
