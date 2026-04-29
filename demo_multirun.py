"""
demo_multirun.py — Mixed-format (multi-run) text slide demo using pptx_functions.

Demonstrates the multi-run "Text" mode: a single paragraph built from multiple
text segments, each with independent font name, size, colour, and style.

The date and status are runtime values — they are built here in Python and
injected into the config before passing it to build_presentation.

Usage
-----
    python demo_multirun.py
"""

import json
from datetime import datetime, timezone
from pptx_functions import build_presentation

with open("config_multirun.json") as f:
    slide_config = json.load(f)

today  = datetime.now(timezone.utc)
status = "non-operational"

slide_config["Details"]["Created"] = today

# Build multi-run text dict and inject into the config.
# Keys are the literal text segments; values are per-run font overrides.
date_str = today.strftime("%d%b%y").upper()

slide_config["Slides"]["Slide 01"]["Objects"]["Status Line"]["Text"] = {
    "(PORTION MARKING) ":  {"Bold?": True,  "Font Name": "Arial Narrow",  "Font Size": 8,    "Font Color": "blue"},
    "Overview: ":          {"Bold?": True,  "Font Name": "Comic Sans MS", "Font Size": 12,   "Font Color": "#252525"},
    "As of ":              {"Bold?": True,  "Font Name": "Comic Sans MS", "Font Size": 12,   "Font Color": "#cb181d"},
    date_str:              {"Bold?": True,  "Font Name": "Comic Sans MS", "Font Size": 16.5, "Font Color": "purple",  "Underline?": True},
    ", the network is  ":  {"Bold?": True,  "Font Name": "Comic Sans MS", "Font Size": 12,   "Font Color": "#969696"},
    status.upper():        {"Bold?": True,  "Font Name": "Bradley Hand",  "Font Size": 18,   "Font Color": "orange",  "Underline?": True},
}

filename = slide_config["Details"]["Filename"] % today.strftime("%Y-%m-%d %H%MZ")
prs = build_presentation(slide_config, verbose=True)
prs.save(filename)
print(f"Saved: {filename}")
