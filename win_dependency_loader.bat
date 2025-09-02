@echo off
REM This code install packages need for Rollout Automation bot and must be installed with admin

:: Look for python
for /f "delims=" %%a in ('where python') do set PYTHON_HOME=%%a
if exist "%PYTHON_HOME%" (
    echo Python location is in: %PYTHON_HOME%
    echo/
 ) else (
    echo Python does not exist in the system, please contact ITBar.
    echo/
    goto :eof
 )


:: Common Package - Webtool
set list=gspread;python-dotenv;selenium;webdriver-manager;google-api-python-client;google-auth;oauth2client;requests;

:: MTPOS Specific
set list=%list%;PyQt5;pywinauto

:: Translate list into array
setlocal enabledelayedexpansion
set "notInstalled="

for %%p in (%list%) do (
    python -m pip show %%p >nul 2>&1
    if errorlevel 1 (
        echo %%p is NOT installed.
        set "notInstalled=!notInstalled! %%p"
    ) else (
        echo %%p is installed. Removing from list.
    )
)

:: Show remaining packages
echo.
echo Remaining packages (to be installed):
echo !notInstalled!
echo.

:: Install packages in system
if defined notInstalled (
   echo Installing missing packages...
   python -m pip install !notInstalled!
   echo Installation Complete
) else (
   echo All packages are already installed.
)
pause
