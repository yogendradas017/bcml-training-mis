@echo off
echo ============================================
echo  BCML Training Management System
echo  Starting local server...
echo ============================================

cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.10+
    pause
    exit /b 1
)

REM Install dependencies if needed
if not exist ".deps_installed" (
    echo Installing required packages...
    pip install -r requirements.txt
    echo. > .deps_installed
)

REM Initialise DB and seed employees on first run
if not exist "data\training.db" (
    echo Setting up database for first time...
    python -c "from app import init_db; init_db()"
    echo Importing employee data from Master Emp Data folder...
    python seed.py
)

echo.
echo System is ready!
echo Open your browser and go to:  http://localhost:5000
echo.
echo Login credentials:
echo   SPOC (any plant):   username = plantname   password = bcml@1234
echo   Central Team:       username = central      password = bcml@1234
echo   Admin:              username = admin        password = admin@bcml
echo.
echo Press Ctrl+C to stop the server.
echo ============================================

python app.py
pause
