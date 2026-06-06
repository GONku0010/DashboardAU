#!/bin/bash
# Ignore broken Streamlit port env vars from Railway.
unset STREAMLIT_SERVER_PORT
unset STREAMLIT_SERVER_ADDRESS
export PORT="${PORT:-8501}"
exec streamlit run app.py --server.address 0.0.0.0 --server.port "$PORT" --logger.level=error
