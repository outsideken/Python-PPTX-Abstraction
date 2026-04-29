"""
demo_bullets.py — Bullet text slide demo using pptx_functions.

Demonstrates the "Bullets" text mode: a list of [text, level] pairs
with per-indent font-size mapping.

Usage
-----
    python demo_bullets.py
"""

import json
from datetime import datetime, timezone
from pptx_functions import build_presentation

with open("config_bullets.json") as f:
    slide_config = json.load(f)

today = datetime.now(timezone.utc)
slide_config["Details"]["Created"] = today

filename = slide_config["Details"]["Filename"] % today.strftime("%Y-%m-%d %H%MZ")
prs = build_presentation(slide_config, verbose=True)
prs.save(filename)
print(f"Saved: {filename}")
