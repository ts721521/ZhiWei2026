#!/bin/bash
cd "$(dirname "$0")"

# Check for virtual environment
if [ -d ".venv" ]; then
    PYTHON_CMD="./.venv/bin/python"
else
    # Fallback to system python/brew python
    if command -v python3 &> /dev/null; then
        PYTHON_CMD="python3"
    else
        echo "Error: Python 3 is not installed or not in PATH."
        exit 1
    fi
fi

echo "Launching Office GUI with $PYTHON_CMD..."
"$PYTHON_CMD" office_gui.py
