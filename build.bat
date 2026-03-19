@echo off
setlocal

cd /d "%~dp0"

set "APP_NAME=Tools Other CE V1"
set "ENTRY=app.py"
set "UPDATER_ENTRY=updater.py"
set "ICON=Iconapp.ico"
set "PY_EXE=.venv\Scripts\python.exe"

if not exist "%PY_EXE%" (
  echo [ERROR] Python not found at "%PY_EXE%"
  echo Create venv first: py -m venv .venv
  exit /b 1
)

if not exist "%ENTRY%" (
  echo [ERROR] Entry file not found: "%ENTRY%"
  exit /b 1
)

if not exist "%UPDATER_ENTRY%" (
  echo [ERROR] Updater file not found: "%UPDATER_ENTRY%"
  exit /b 1
)

if not exist "%ICON%" (
  echo [ERROR] Icon file not found: "%ICON%"
  exit /b 1
)

echo [1/3] Checking PyInstaller...
"%PY_EXE%" -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
  echo Installing PyInstaller...
  "%PY_EXE%" -m pip install pyinstaller
  if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller.
    exit /b 1
  )
)

echo [2/3] Building EXE...
"%PY_EXE%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name "%APP_NAME%" ^
  --icon "%ICON%" ^
  --add-data "%ICON%;." ^
  --collect-all pyreadstat ^
  --hidden-import pyreadstat._readstat_parser ^
  --hidden-import pyreadstat._readstat_writer ^
  "%ENTRY%"

if errorlevel 1 (
  echo [ERROR] Main build failed.
  exit /b 1
)

echo [3/4] Building Updater...
"%PY_EXE%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name "Tools Other CE Updater" ^
  --icon "%ICON%" ^
  --add-data "%ICON%;." ^
  "%UPDATER_ENTRY%"

if errorlevel 1 (
  echo [ERROR] Updater build failed.
  exit /b 1
)

echo [4/4] Done.
echo EXE: "%cd%\dist\%APP_NAME%.exe"
echo UPDATER: "%cd%\dist\Tools Other CE Updater.exe"
exit /b 0
