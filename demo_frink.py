"""
demo_frink.py — Professor Frink single-image slide demo using pptx_functions.

Loads the Professor Frink configuration JSON and generates a .pptx.

Usage
-----
    python demo_frink.py
"""

import json
from datetime import datetime, timezone
from pptx_functions import build_presentation

with open("Professor Frink Configuration (2025-03-31).json") as f:
    slide_config = json.load(f)

today = datetime.now(timezone.utc)
slide_config["Details"]["Created"] = today

filename = slide_config["Details"]["Filename"] % today.strftime("%Y-%m-%d %H%MZ")
prs = build_presentation(slide_config, verbose=True)
prs.save(filename)
print(f"Saved: {filename}")
