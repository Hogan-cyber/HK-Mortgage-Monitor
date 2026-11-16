@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
pushd "%~dp0"
if errorlevel 1 (
  echo Cannot open this folder. Extract the ZIP to a writable local folder first.
  pause
  exit /b 1
)
py -3 -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>&1
if not errorlevel 1 goto run_py
python -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>&1
if not errorlevel 1 goto run_python
python3 -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>&1
if not errorlevel 1 goto run_python3
for %%P in ("%USERPROFILE%\anaconda3\python.exe" "%USERPROFILE%\miniconda3\python.exe" "%LOCALAPPDATA%\anaconda3\python.exe" "%LOCALAPPDATA%\miniconda3\python.exe" "%ProgramData%\anaconda3\python.exe" "%ProgramData%\miniconda3\python.exe") do (
  if exist "%%~P" (
    "%%~P" -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>&1
    if not errorlevel 1 (
      set "MORTGAGE_PY=%%~P"
      goto run_absolute
    )
  )
)
echo Updating requires Python 3.10 or newer. Install Python from python.org, then run this file again.
echo You can still open HK_Mortgage_Dashboard.html without Python or internet.
set "MORTGAGE_EXIT=1"
goto finish
:run_py
py -3 "%~dp0update_mortgage_data.py" %*
set "MORTGAGE_EXIT=%errorlevel%"
goto finish
:run_python
python "%~dp0update_mortgage_data.py" %*
set "MORTGAGE_EXIT=%errorlevel%"
goto finish
:run_python3
python3 "%~dp0update_mortgage_data.py" %*
set "MORTGAGE_EXIT=%errorlevel%"
goto finish
:run_absolute
"%MORTGAGE_PY%" "%~dp0update_mortgage_data.py" %*
set "MORTGAGE_EXIT=%errorlevel%"
:finish
if "%MORTGAGE_EXIT%"=="0" (
  echo Completed. Open HK_Mortgage_Dashboard.html or reload its browser tab.
) else (
  echo Update did not finish. See the message above and update_status.txt if present.
)
popd
pause
exit /b %MORTGAGE_EXIT%
