@echo off
setlocal

rem Build a one-file Windows executable into the scripts folder.
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "PROJECT_ROOT=%SCRIPT_DIR%\.."
set "PY_EXE=C:\Users\wayne\AppData\Local\Python\pythoncore-3.14-64\python.exe"

rem Clear environment overrides that can hide site-packages.
set "PYTHONHOME="
set "PYTHONPATH="

pushd "%PROJECT_ROOT%"
"%PY_EXE%" -m PyInstaller --clean --distpath "%SCRIPT_DIR%" --workpath "%SCRIPT_DIR%\build" ^
  "%PROJECT_ROOT%\qamRoster.spec"
popd

endlocal
