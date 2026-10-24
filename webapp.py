"""Data Trust Dashboard launcher: python webapp.py  (or double-click TrustScore.exe)."""
import socket
import threading
import webbrowser

from trust_app import create_app


def free_port(start=5000):
    for port in range(start, start + 50):
        with socket.socket() as probe:
            if probe.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start


if __name__ == "__main__":
    port = free_port()
    url = f"http://127.0.0.1:{port}"
    print(f"Data Trust Dashboard is running at {url}")
    print("Keep this window open; close it to stop the app.")
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    create_app().run(host="127.0.0.1", port=port, debug=False)
