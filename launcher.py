
import os
import sys
import socket
import threading
import time
import webbrowser
import traceback
from datetime import datetime
from pathlib import Path

os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

from streamlit.web import cli as stcli


def get_app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
os.chdir(APP_DIR)

PORT = 8501
URL = f"http://127.0.0.1:{PORT}"
LOG_FILE = APP_DIR / "launcher.log"

browser_opened = threading.Event()
stop_monitor = threading.Event()


def write_log(message):
    try:
        with LOG_FILE.open("a", encoding="utf-8") as log:
            log.write(
                f"[{datetime.now():%Y-%m-%d %H:%M:%S}] "
                f"{message}\n"
            )
    except OSError:
        pass


def server_is_ready():
    try:
        with socket.create_connection(
            ("127.0.0.1", PORT),
            timeout=1
        ):
            return True
    except OSError:
        return False


def open_browser_when_ready():
    while not stop_monitor.is_set():
        if server_is_ready():
            if not browser_opened.is_set():
                webbrowser.open(URL)
                browser_opened.set()
                write_log("Browser opened.")
            return

        time.sleep(1)


def run_streamlit(app_path):
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.address=127.0.0.1",
        f"--server.port={PORT}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--server.fileWatcherType=none",
    ]

    return stcli.main()


def main():
    app_path = APP_DIR / "streamlit_app.py"

    if not app_path.exists():
        message = f"Application file not found: {app_path}"
        write_log(message)
        raise FileNotFoundError(message)

    write_log("LYNX Cross Search launcher started.")

    if server_is_ready():
        write_log(
            f"Port {PORT} is already in use. "
            "Another application instance may be running."
        )
        webbrowser.open(URL)
        return

    threading.Thread(
        target=open_browser_when_ready,
        daemon=True
    ).start()

    restart_count = 0

    try:
        while True:
            try:
                write_log("Starting Streamlit.")

                result = run_streamlit(app_path)

                write_log(
                    f"Streamlit exited with result: {result}"
                )

            except KeyboardInterrupt:
                write_log("Application stopped by user.")
                break

            except SystemExit as error:
                write_log(
                    f"Streamlit SystemExit: {error.code}"
                )

            except Exception:
                write_log(
                    "Streamlit crashed:\n"
                    + traceback.format_exc()
                )

            restart_count += 1

            write_log(
                f"Restart attempt #{restart_count} in 5 seconds."
            )

            time.sleep(5)

    finally:
        stop_monitor.set()
        write_log("Launcher stopped.")


if __name__ == "__main__":
    main()
