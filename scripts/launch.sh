#!/usr/bin/env bash
set -e
cd "$(dirname "$0")/.."

echo "Installing required libraries..."
if command -v py &> /dev/null; then
    py -m pip install -r requirements.txt
elif command -v python3 &> /dev/null; then
    python3 -m pip install -r requirements.txt
elif command -v python &> /dev/null; then
    python -m pip install -r requirements.txt
else
    echo "Error: No Python interpreter found. Install Python 3.8+."
    exit 1
fi

echo "Starting CV Manager..."
if command -v py &> /dev/null; then
    py src/main.py "$@"
elif command -v python3 &> /dev/null; then
    python3 src/main.py "$@"
elif command -v python &> /dev/null; then
    python src/main.py "$@"
fi
