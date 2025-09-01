@echo off

setlocal ENABLEDELAYEDEXPANSION



set "PID_FILE=%TEMP%\xlsx_selector_pids.txt"



echo.

echo ====================================

echo   Starting the xlsxSelector Tool

echo ====================================

echo.



:: Change directory to the location of this script

cd /d "%~dp0"

echo Changed directory to: %cd%

echo.



:: Check for Python

where python >nul 2>nul

if %errorlevel% neq 0 (

    echo Error: Python not found.

    echo Please install Python from https://www.python.org/

    echo Make sure to check "Add Python to PATH" during installation.

    pause

    exit /b 1

)



:: Install dependencies from requirements.txt

echo Installing dependencies from requirements.txt...

python -m pip install -r requirements.txt

if %errorlevel% neq 0 (

    echo.

    echo Error: Failed to install dependencies.

    echo Please check your internet connection or run "pip install -r requirements.txt" manually.

    echo.

    pause

    exit /b 1

)

echo Dependencies installed successfully.



:: Clean up old PID file

if exist "%PID_FILE%" del "%PID_FILE%"



:: Start Backend Server and get its PID

echo.

echo Starting backend server on port 5000...

start "Backend Server" /min python backend\app.py

for /f "tokens=2" %%i in ('tasklist /nh /fi "imagename eq python.exe" /fi "windowtitle eq Backend Server"') do (

    echo %%i >> "%PID_FILE%"

    goto :start_frontend

)



:start_frontend

:: Start Frontend Server and get its PID

echo.

echo Starting frontend server on port 8000...

start "Frontend Server" /min python -m http.server 8000 --directory frontend

for /f "tokens=2" %%i in ('tasklist /nh /fi "imagename eq python.exe" /fi "windowtitle eq Frontend Server"') do (

    echo %%i >> "%PID_FILE%"

    goto :browser_and_prompt

)



:browser_and_prompt

:: Wait for a few seconds to let the servers start

timeout /t 5 > nul



:: Open the web browser

echo.

echo Opening the web browser. Please go to http://127.0.0.1:8000

start http://127.0.0.1:8000



echo.

echo ------------------------------------

echo All services have started.

echo Please keep the 'Backend Server' and 'Frontend Server' windows open until you are finished.

echo.

echo To close all services, press Ctrl+C and then "Y".

echo.

pause >nul



:: Stop the services on Ctrl+C

:stop_services

echo Stopping services...

for /f %%i in ('type "%PID_FILE%"') do (

    taskkill /PID %%i /F >nul 2>nul

)

if exist "%PID_FILE%" del "%PID_FILE%"

echo Services stopped.

pause