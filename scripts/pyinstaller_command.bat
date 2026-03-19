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
"%PY_EXE%" -m PyInstaller --onefile --name "qamRoster" --distpath "%SCRIPT_DIR%" --workpath "%SCRIPT_DIR%\build" --specpath "%SCRIPT_DIR%" ^
  --add-data "%PROJECT_ROOT%\modules\credentials.json;modules" ^
  --add-data "%PROJECT_ROOT%\project.json;." ^
  --add-data "%PROJECT_ROOT%\QAM-Logo-1-2048x1310whiteBGRND.png;." ^
  "%PROJECT_ROOT%\main.py"
popd

endlocal
