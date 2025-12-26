#!/usr/bin/env bash
cd /app

echo "Searching for uvicorn"
find / -type f -name "uvicorn" 2>/dev/null


source /app/venv/bin/activate
which python
which uvicorn



# Use venv python
exec uvicorn main:app --host 0.0.0.0 --port 8000

