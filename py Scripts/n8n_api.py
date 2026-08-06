#!/usr/bin/env python3
"""
n8n_api.py — HTTP API wrapper for N8N

Exposes the Layer 1 router (and future layers) as REST endpoints
so N8N can call them via HTTP Request nodes.

Start the server:
    pip install flask
    python n8n_api.py
    # or: python n8n_api.py --port 5055

N8N HTTP Request node settings:
    Method: POST
    URL:    http://localhost:5055/detect
    Body:   {"url": "https://www.webnovel.com/book/..."}

Endpoints:
    POST /detect
        Body: { "url": "...", "start_chapter": 1, "end_chapter": null, "search_translations": true }
        Returns: full Layer 1 JSON payload

    GET  /platforms
        Returns: list of supported platforms

    GET  /health
        Returns: { "status": "ok" }
"""

import argparse
import json
import logging
import sys
import os

# Add this script's directory to sys.path so layer1_router can be imported
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, jsonify
from layer1_router import route, PLATFORMS

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("n8n_api")

app = Flask(__name__)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/platforms", methods=["GET"])
def list_platforms():
    """Return all supported platforms and their configs."""
    return jsonify({
        "count": len(PLATFORMS),
        "platforms": [
            {
                "domain": domain,
                "platform": cfg["platform"],
                "source_language": cfg["source_language"],
                "needs_translation": cfg["needs_translation"],
                "scraper_script": cfg.get("scraper_script"),
                "needs_browser": cfg.get("needs_browser", True),
                "needs_login": cfg.get("needs_login", False),
            }
            for domain, cfg in PLATFORMS.items()
        ],
    })


@app.route("/detect", methods=["POST"])
def detect():
    """
    Layer 1: Detect platform, extract metadata, find translations.

    Request body (JSON):
        {
            "url": "https://www.webnovel.com/book/...",
            "start_chapter": 1,          // optional
            "end_chapter": null,          // optional, null = auto
            "search_translations": true   // optional, default true
        }

    Response: full Layer 1 JSON payload (see layer1_router.route())
    """
    body = request.get_json(silent=True) or {}
    url = body.get("url", "").strip()

    if not url:
        return jsonify({"status": "error", "error": "Missing 'url' in request body"}), 400

    start_chapter = int(body.get("start_chapter", 1))
    end_chapter = body.get("end_chapter")
    if end_chapter is not None:
        end_chapter = int(end_chapter)
    search_translations = bool(body.get("search_translations", True))

    logger.info("detect request: url=%s start=%d end=%s", url, start_chapter, end_chapter)

    try:
        result = route(
            url=url,
            start_chapter=start_chapter,
            end_chapter=end_chapter,
            search_translations=search_translations,
        )
        status_code = 200 if result.get("status") == "ok" else 422
        return jsonify(result), status_code
    except Exception as e:
        logger.exception("Error in /detect: %s", e)
        return jsonify({"status": "error", "error": str(e)}), 500


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="N8N API server for content scraper")
    parser.add_argument("--host", default="0.0.0.0", help="Bind host (default: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=5055, help="Port (default: 5055)")
    parser.add_argument("--debug", action="store_true", help="Enable Flask debug mode")
    args = parser.parse_args()

    logger.info("Starting N8N API on %s:%d", args.host, args.port)
    logger.info("Endpoints: POST /detect | GET /platforms | GET /health")
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
