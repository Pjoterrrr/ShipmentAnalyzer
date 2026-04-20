from __future__ import annotations

import socket
import sys
import threading
import webbrowser
from pathlib import Path

from streamlit.web import bootstrap


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def find_free_port(preferred: int = 8501) -> int:
    for port in range(preferred, preferred + 30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return preferred


def main() -> None:
    base_dir = app_base_dir()
    app_script = base_dir / "app.py"
    port = find_free_port()
    address = "127.0.0.1"
    target_url = f"http://{address}:{port}"

    threading.Timer(1.6, lambda: webbrowser.open(target_url)).start()

    bootstrap.run(
        str(app_script),
        False,
        [],
        {
            "server.headless": True,
            "server.port": port,
            "server.address": address,
            "browser.gatherUsageStats": False,
            "server.fileWatcherType": "none",
        },
    )


if __name__ == "__main__":
    main()
