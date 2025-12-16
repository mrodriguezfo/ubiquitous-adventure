#!/usr/bin/env python3
"""
run_app.py

Start the FastAPI app with a single python command:

  python run_app.py

Environment variables:
  PORT    - port to listen on (default 12000)
  HOST    - host to bind to (default 0.0.0.0)
  RELOAD  - if set to 1/true, uvicorn reload is enabled (for dev)

This script ensures the repository root is on sys.path so imports like
`from src.processing import ...` work when the app is started.
"""
import os
import sys
from pathlib import Path

# Add repo root to sys.path so package imports like `src.*` resolve
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import uvicorn

if __name__ == '__main__':
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', '12000'))
    reload = os.environ.get('RELOAD', '').lower() in ('1', 'true', 'yes')

    # Run the app defined in src.app_main:app
    uvicorn.run('src.app_main:app', host=host, port=port, reload=reload)
