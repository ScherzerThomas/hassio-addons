#!/usr/bin/env bash
cd /app

# Use venv python
/app/venv/bin/uvicorn main:app --host 0.0.0.0 --port 8000
