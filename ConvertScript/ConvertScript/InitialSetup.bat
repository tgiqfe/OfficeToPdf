@echo off
pushd %~dp0

echo # Create venv
python -m venv venv

echo # upgrade pip
venv\Scripts\python.exe -m pip install --upgrade pip

echo # install pywin32
venv\Scripts\python.exe -m pip install pywin32
