#!/usr/bin/env bash
cd /app

source /app/venv/bin/activate

# Use venv python
exec uvicorn main:app --host 0.0.0.0 --port 8000
