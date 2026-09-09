"""
Production entry point. The Flask dev server (`python purchase_scheduler.py`)
is fine for testing but is single-threaded and not meant for long-running
production use. Run this instead:

    python run_waitress.py

It starts the same APScheduler background job and serves the Flask app
through waitress, a production-grade pure-Python WSGI server (no extra
OS packages needed, works on Windows and Linux).
"""

from waitress import serve
from purchase_scheduler import app, start_scheduler, APP_HOST, APP_PORT, scheduler, logger

if __name__ == "__main__":
    start_scheduler()
    try:
        logger.info("Serving on %s:%s via waitress", APP_HOST, APP_PORT)
        serve(app, host=APP_HOST, port=APP_PORT)
    finally:
        scheduler.shutdown(wait=False)
