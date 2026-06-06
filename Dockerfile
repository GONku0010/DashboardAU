FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt ./
RUN python -m pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501

CMD ["bash", "-lc", "unset STREAMLIT_SERVER_PORT STREAMLIT_SERVER_ADDRESS && exec python -m streamlit run app.py --server.address 0.0.0.0 --server.port \"${PORT:-8501}\" --logger.level=error"]
