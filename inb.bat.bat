@echo off
setlocal

REM ---------------------------------------------------------------------------
REM  Inbound Weekly Update - load vendor exports and stage them into the template.
REM
REM  All paths resolve from %~dp0 (the folder holding this .bat), so this works
REM  for any user and from any working directory - double-click it, or call it
REM  from Task Scheduler.
REM ---------------------------------------------------------------------------

set "REPO=%~dp0"
set "VENV=%REPO%.venv"
set "PY=%VENV%\Scripts\python.exe"

if not exist "%PY%" (
    echo [ERROR] Virtual environment not found at:
    echo         %VENV%
    echo.
    echo Create it with:
    echo     python -m venv "%VENV%"
    echo     "%PY%" -m pip install -r "%REPO%requirements.txt"
    echo.
    pause
    exit /b 1
)

REM Activate the environment.
call "%VENV%\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] Could not activate the virtual environment.
    pause
    exit /b 1
)

REM Run from the repo so relative paths behave.
cd /d "%REPO%"

echo Using: %PY%
echo Running mainfinal4.py ...
echo.

REM Invoked by full path rather than bare "python" so the correct interpreter is
REM used even if activation left PATH in an unexpected state.
"%PY%" "%REPO%mainfinal4.py"
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
    echo [ERROR] Script exited with code %RC%
) else (
    echo Finished successfully.
)

REM Keeps the window open when double-clicked. Remove this line for Task Scheduler.
pause
exit /b %RC%
