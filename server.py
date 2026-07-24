#!/usr/bin/env python3
"""
Flask server — serves jobs from jobs.json and hosts the web UI.
Locally: also exposes /api/refresh to re-run the scraper.
"""
import json
import os
import subprocess
import sys
import threading
from pathlib import Path

from flask import Flask, jsonify, send_from_directory, request

BASE      = Path(__file__).parent
JOBS_FILE = BASE / "jobs.json"

app = Flask(__name__, static_folder=str(BASE / "static"))


def read_jobs():
    if not JOBS_FILE.exists():
        return {"jobs": [], "total": 0, "platforms": [], "scraped_at": None}
    return json.loads(JOBS_FILE.read_text())


@app.route("/api/jobs")
def api_jobs():
    return jsonify(read_jobs())


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    # Only allow from localhost — scraper can't run on Render anyway
    if request.remote_addr not in ("127.0.0.1", "::1"):
        return jsonify({"error": "not allowed"}), 403
    def run():
        subprocess.run([sys.executable, str(BASE / "find_jobs.py")], cwd=str(BASE))
    threading.Thread(target=run, daemon=True).start()
    return jsonify({"status": "started"})


@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    static = BASE / "static"
    target = static / path
    if path and target.exists():
        return send_from_directory(str(static), path)
    return send_from_directory(str(static), "index.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print(f"\n  Job Finder Web UI")
    print(f"  Open → http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
