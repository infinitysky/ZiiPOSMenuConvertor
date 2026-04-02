"""
ZiiPOS Menu Converter V3 -- Entry Point
pywebview window + Flask backend server
"""
import sys
import os
import socket
import threading
import webview

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

from app import create_app


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def start_server(app, port):
    app.run(host="127.0.0.1", port=port, threaded=True, use_reloader=False)


if __name__ == "__main__":
    port = find_free_port()
    app = create_app()

    t = threading.Thread(target=start_server, args=(app, port), daemon=True)
    t.start()

    webview.create_window(
        "ZiiPOS Menu Converter V3",
        f"http://127.0.0.1:{port}",
        width=960,
        height=780,
        resizable=True,
        min_size=(800, 600),
    )
    webview.start()
