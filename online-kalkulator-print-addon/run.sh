#!/usr/bin/env bash
cd /app

# Start LibreOffice headless listener
libreoffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" &

# Start FastAPI app
uvicorn main:app --host 0.0.0.0 --port 8000
