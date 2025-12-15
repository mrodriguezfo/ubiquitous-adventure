#!/bin/bash
# Run the FastAPI app on the provided runtime port
export PYTHONPATH="${PYTHONPATH}:$(pwd)/src"
uvicorn src.app_main:app --host 0.0.0.0 --port 12000 --loop auto
