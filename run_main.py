import ctypes
import os
import socket
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

def resolve_path(path):
    if getattr(sys, "frozen", False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(__file__)
    return os.path.join(basedir, path)

def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]

def open_browser_when_ready(url, port):
    for _ in range(80):
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.3):
                webbrowser.open(url)
                return
        except OSError:
            time.sleep(0.25)
    webbrowser.open(url)

def log_path():
    base = os.environ.get("LOCALAPPDATA") or os.path.dirname(os.path.abspath(sys.argv[0]))
    target = Path(base) / "DutyScheduler"
    try:
        target.mkdir(parents=True, exist_ok=True)
    except OSError:
        target = Path(os.path.dirname(os.path.abspath(sys.argv[0]))) / "logs"
        target.mkdir(parents=True, exist_ok=True)
    return target / "DutyScheduler.log"

def write_log(details):
    path = log_path()
    path.write_text(details, encoding="utf-8")
    return path

def show_error(message):
    try:
        ctypes.windll.user32.MessageBoxW(0, message, "Duty Scheduler", 0x10)
    except Exception:
        pass

def pause_console(message):
    print()
    print(message)
    print()
    if sys.stdin and sys.stdin.isatty():
        try:
            input("Press Enter to close this window...")
        except EOFError:
            pass

def main():
    import streamlit.web.cli as stcli

    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(os.path.abspath(sys.executable)))
    else:
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

    if os.environ.get("DUTY_SCHEDULER_IMPORT_CHECK") == "1":
        import logic  # noqa: F401
        import openpyxl  # noqa: F401
        import pandas  # noqa: F401
        import pdfplumber  # noqa: F401
        import streamlit  # noqa: F401
        import xlsxwriter  # noqa: F401

        app_file = Path(resolve_path("duty_app.py"))
        if not app_file.exists():
            raise FileNotFoundError(f"Missing app entry: {app_file}")
        print("IMPORT_CHECK_OK", flush=True)
        return 0

    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    threading.Thread(target=open_browser_when_ready, args=(url, port), daemon=True).start()

    sys.argv = [
        "streamlit",
        "run",
        resolve_path("duty_app.py"),
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.fileWatcherType=none",
        "--browser.gatherUsageStats=false",
        f"--server.port={port}",
        "--server.address=127.0.0.1",
    ]
    return stcli.main()

if __name__ == "__main__":
    try:
        exit_code = main()
        if os.environ.get("DUTY_SCHEDULER_IMPORT_CHECK") == "1":
            sys.exit(exit_code or 0)
        if exit_code not in (None, 0):
            raise RuntimeError(f"Streamlit exited with code {exit_code}")
        pause_console("Duty Scheduler has exited.")
        sys.exit(exit_code or 0)
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
        if code in (None, 0):
            if os.environ.get("DUTY_SCHEDULER_IMPORT_CHECK") != "1":
                pause_console("Duty Scheduler has exited.")
            raise

        details = f"SystemExit: {exc.code}\n\n{traceback.format_exc()}"
        path = write_log(details)
        print(details)
        show_error(f"Duty Scheduler failed to start.\n\nLog file:\n{path}")
        pause_console(f"Duty Scheduler failed to start. Log file: {path}")
        sys.exit(code)
    except BaseException:
        details = traceback.format_exc()
        path = write_log(details)
        print(details)
        show_error(f"Duty Scheduler failed to start.\n\nLog file:\n{path}")
        pause_console(f"Duty Scheduler failed to start. Log file: {path}")
        sys.exit(1)
