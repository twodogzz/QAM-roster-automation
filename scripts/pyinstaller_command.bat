@echo off
setlocal

rem Build a one-file Windows executable into the scripts folder.
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "PROJECT_ROOT=%SCRIPT_DIR%\.."
set "PY_EXE=C:\Users\wayne\AppData\Local\Python\pythoncore-3.14-64\python.exe"
set "ICON_SOURCE=%PROJECT_ROOT%\QAM-Logo-1-2048x1310whiteBGRND.png"
set "ICON_OUTPUT=%PROJECT_ROOT%\QAM-Logo-1-2048x1310whiteBGRND.ico"
set "ICON_SCRIPT=%SCRIPT_DIR%\generate_windows_icon.ps1"

rem Clear environment overrides that can hide site-packages.
set "PYTHONHOME="
set "PYTHONPATH="

if not exist "%ICON_SOURCE%" (
  echo Icon source not found: "%ICON_SOURCE%"
  exit /b 1
)

if not exist "%ICON_SCRIPT%" (
  echo Icon generation script not found: "%ICON_SCRIPT%"
  exit /b 1
)

echo Generating Windows icon...
powershell -ExecutionPolicy Bypass -File "%ICON_SCRIPT%" -SourcePng "%ICON_SOURCE%" -DestinationIco "%ICON_OUTPUT%"
if errorlevel 1 (
  echo Failed to generate Windows icon.
  exit /b 1
)

pushd "%PROJECT_ROOT%"
"%PY_EXE%" -m PyInstaller --clean --distpath "%SCRIPT_DIR%" --workpath "%SCRIPT_DIR%\build" ^
  "%PROJECT_ROOT%\qamRoster.spec"
if errorlevel 1 (
  popd
  exit /b 1
)
popd

endlocal
