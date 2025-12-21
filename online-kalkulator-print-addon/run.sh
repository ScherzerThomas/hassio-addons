#!/usr/bin/env bash
cd /app

# Start FastAPI
uvicorn main:app --host 0.0.0.0 --port 8000
