# -*- coding: utf-8 -*-
"""
Local sync server for MyFinance.

Serves the built site (06_Web) AND a tiny JSON sync endpoint so a phone on the
same Wi-Fi can read/write the SAME data as this PC. Fully local: no internet,
no cloud, no third-party host. Data lives in 03_System/sync_data.json (NTFS-locked).

Run:  python sync_server.py
Then open on the phone:  http://<this-PC-LAN-IP>:8765/
"""
import json
import os
import socket
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.normpath(os.path.join(ROOT, "..", "06_Web"))
DATA = os.path.normpath(os.path.join(ROOT, "..", "03_System", "sync_data.json"))
PORT = int(os.environ.get("MF_SYNC_PORT", "8765"))

MIME = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".ico": "image/x-icon",
    ".woff2": "font/woff2",
    ".woff": "font/woff",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".csv": "text/csv; charset=utf-8",
    ".txt": "text/plain; charset=utf-8",
}

SECTIONS = ("log", "cfg", "base", "flags", "ai")


def load():
    try:
        with open(DATA, "r", encoding="utf-8") as f:
            d = json.load(f)
            if isinstance(d, dict):
                return d
    except Exception:
        pass
    return {"v": 1}


def save(d):
    os.makedirs(os.path.dirname(DATA), exist_ok=True)
    tmp = DATA + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False)
    os.replace(tmp, DATA)


def merge(server, client):
    """Per-section last-write-wins by timestamp."""
    out = {"v": 1}
    for k in SECTIONS:
        s = server.get(k) or {}
        c = client.get(k) or {}
        try:
            if (c.get("ts") or 0) >= (s.get("ts") or 0):
                out[k] = c
            else:
                out[k] = s
        except Exception:
            out[k] = s
    return out


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _json(self, code, obj):
        b = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors()
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        path = self.path.split("?")[0]
        if path.startswith("/__sync_ok"):
            return self._json(200, {"ok": True})
        if path.startswith("/__sync"):
            return self._json(200, load())
        return self._serve(path)

    def do_POST(self):
        path = self.path.split("?")[0]
        if path.startswith("/__sync"):
            try:
                n = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(max(0, n))
                client = json.loads(body.decode("utf-8"))
            except Exception:
                return self._json(400, {"error": "bad json"})
            merged = merge(load(), client)
            save(merged)
            return self._json(200, merged)
        self.send_response(404)
        self.end_headers()

    def _serve(self, path):
        rel = path.lstrip("/") or "index.html"
        fp = os.path.normpath(os.path.join(WEB, rel))
        if not fp.startswith(WEB) or not os.path.isfile(fp):
            self.send_response(404)
            self.end_headers()
            return
        ext = os.path.splitext(fp)[1].lower()
        try:
            with open(fp, "rb") as f:
                data = f.read()
        except Exception:
            self.send_response(500)
            self.end_headers()
            return
        self.send_response(200)
        self.send_header("Content-Type", MIME.get(ext, "application/octet-stream"))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("192.168.1.1", 80))
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()


def main():
    ip = lan_ip()
    print("MyFinance local sync server")
    print("  this PC : http://127.0.0.1:%d/" % PORT)
    print("  phone   : http://%s:%d/   (same Wi-Fi)" % (ip, PORT))
    print("  data    : %s" % DATA)
    print("  stop    : Ctrl+C")
    try:
        ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\nstopped")
        sys.exit(0)


if __name__ == "__main__":
    main()
