"""Start the integrated Waitress app and its E&I deadline evaluator."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from integrated_app import application
from core.workflow_worker import start_worker
from waitress import serve
from deployment.server_config import server_address, waitress_proxy_options


if __name__ == '__main__':
    stop = start_worker()
    try:
        host, port = server_address()
        serve(application, host=host, port=port, threads=8, **waitress_proxy_options())
    finally:
        stop.set()
