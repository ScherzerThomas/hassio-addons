#!/usr/bin/env bash
set -e
cd /app

echo "Activating virtual environment..."
source /app/venv/bin/activate

echo "Python interpreter:"
which python

echo "Uvicorn executable:"
which uvicorn

# Debug mode (used when running locally with VS Code)
if [ "$DEBUG" = "1" ]; then
    echo "Starting in DEBUG mode..."
    exec python -m debugpy --listen 0.0.0.0:5678 --wait-for-client \
        -m uvicorn main:app --host 0.0.0.0 --port 8000
fi

# Normal mode (used by Home Assistant OS)
echo "Starting normally..."
exec uvicorn main:app --host 0.0.0.0 --port 8000
