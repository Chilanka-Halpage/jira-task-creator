#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if [ ! -d venv ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

source venv/bin/activate
echo "Installing/updating dependencies..."
pip install -r requirements.txt --quiet

# Local single-user run: remember Jira credentials & mapping across restarts
# by saving them to ./data on disk. Never set this for a shared/public
# deployment (e.g. Streamlit Community Cloud) — it would let one visitor's
# saved Jira token and mapping be seen and overwritten by every other
# visitor, since that disk is shared by everyone who opens the app there.
export JIRA_APP_PERSIST_LOCAL=1
export JIRA_APP_ALLOW_STOP=1

echo "Starting app... a browser tab will open automatically."
streamlit run app.py
