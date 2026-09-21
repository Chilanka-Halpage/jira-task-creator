@echo off
cd /d "%~dp0"

if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)

call venv\Scripts\activate.bat

echo Installing/updating dependencies...
pip install -r requirements.txt --quiet

rem Local single-user run: remember Jira credentials & mapping across restarts
rem by saving them to ./data on disk. Never set this for a shared/public
rem deployment (e.g. Streamlit Community Cloud) — it would let one visitor's
rem saved Jira token and mapping be seen and overwritten by every other
rem visitor, since that disk is shared by everyone who opens the app there.
set JIRA_APP_PERSIST_LOCAL=1
set JIRA_APP_ALLOW_STOP=1

echo Starting app... a browser tab will open automatically.
streamlit run app.py

pause
