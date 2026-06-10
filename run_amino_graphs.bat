@echo off
setlocal
cd /d "%~dp0"

echo ===============================================
echo Amino Eggsactly Graphs - Streamlit Launcher
echo ===============================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creating local Python virtual environment...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo Could not create virtual environment. Please make sure Python 3 is installed.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"

echo Installing / checking Python packages...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo Package installation failed.
    pause
    exit /b 1
)

echo.
echo Starting Streamlit app...
echo Your browser should open automatically.
echo If not, copy the local URL shown below.
echo.
streamlit run app.py

pause
