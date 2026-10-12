import os
os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
import sys
import socket
import threading
import time
import webbrowser
from pathlib import Path

from streamlit.web import cli as stcli


def get_app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
os.chdir(APP_DIR)


def open_browser():
    for _ in range(60):
        try:
            with socket.create_connection(("127.0.0.1", 8501), timeout=1):
                webbrowser.open("http://localhost:8501")
                return
        except OSError:
            time.sleep(1)


if __name__ == "__main__":
    app_path = APP_DIR / "streamlit_app.py"

    if not app_path.exists():
        raise FileNotFoundError(
            f"Application file not found: {app_path}"
        )

    threading.Thread(
        target=open_browser,
        daemon=True
    ).start()

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.address=127.0.0.1",
        "--server.port=8501",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--server.fileWatcherType=none",
    ]

    sys.exit(stcli.main())