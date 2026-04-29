"""
demo_images.py — Image placement slide demo using pptx_functions.

Demonstrates smart aspect-ratio handling (fit: width / height / native),
multi-slide configs, and connector usage.

Slide 1: Single character with the Planet Express ship logo.
Slide 2: Full Planet Express crew — four characters side by side.

Usage
-----
    python demo_images.py
"""

import json
from datetime import datetime, timezone
from pptx_functions import build_presentation

with open("config_images.json") as f:
    slide_config = json.load(f)

today = datetime.now(timezone.utc)
slide_config["Details"]["Created"] = today

filename = slide_config["Details"]["Filename"] % today.strftime("%Y-%m-%d %H%MZ")
prs = build_presentation(slide_config, verbose=True)
prs.save(filename)
print(f"Saved: {filename}")
