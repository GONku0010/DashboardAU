#!/bin/bash
export STREAMLIT_SERVER_PORT=${PORT:-8501}
export STREAMLIT_SERVER_ADDRESS=0.0.0.0
export STREAMLIT_SERVER_HEADLESS=true
export STREAMLIT_CLIENT_TOOLBAR_MODE=minimal
exec streamlit run app.py
