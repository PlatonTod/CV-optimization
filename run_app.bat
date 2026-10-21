@echo off
cd /d "%~dp0"
echo Installing / checking packages...
python -m pip install -q -r requirements.txt
echo Starting Dimax HR...
python -m streamlit run app.py
pause
