# pptx_functions

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![python-pptx](https://img.shields.io/badge/python--pptx-0.6%2B-orange)
![License](https://img.shields.io/badge/license-MIT-green)

A Python abstraction layer for programmatic PowerPoint slide generation using
[python-pptx](https://python-pptx.readthedocs.io/en/latest/index.html).

---

## The Problem with Manual Slides

Building PowerPoint slides by hand is slow, inconsistent, and impossible to
automate.  Copying formatting between slides, adjusting positions pixel by
pixel, and regenerating decks every time the data changes — all of it is
time that should be spent on analysis, not slide assembly.

**pptx_functions** replaces that workflow with a single Python dictionary.
Define your slide content and layout in a `slide_config` dict, call one
function per slide object, and get a pixel-perfect `.pptx` file every time.
The same config can be regenerated with updated data in seconds.

---

## What It Does

```
slide_config dict  →  pptx_functions  →  .pptx file
```

Each slide and each object on it is described by a plain Python dictionary.
`pptx_functions` reads those dictionaries and calls the appropriate
`python-pptx` API, handling coordinate conversion, colour parsing,
aspect-ratio maths, and font formatting so you don't have to.

**Supported slide objects:**

| Object Type | Key | Description |
|---|---|---|
| Text box | `"Text"` | Plain text, mixed-format runs, or bulleted lists |
| Image | `"Image"` | Smart aspect-ratio handling (`fit: width / height / native`) |
| Connector | `"Connector"` | Straight, elbow, or curved lines with dash styling |
| Auto Shape | `"AutoShape"` | Any of 35+ MSO shapes with fill and border formatting |
| Table | `"Table"` | Header row + data rows with per-cell font formatting |
| Banner | `"Banner"` | Top and bottom classification/marking banners |
| Header | `"Header"` | Title text box + separator line + left/right seal images |

---

## Installation

```bash
pip install python-pptx Pillow webcolors
```

Copy `pptx_functions.py` into your project folder (or anywhere on your
Python path), then import it:

```python
from pptx_functions import *        # all functions available directly
import pptx_functions               # also available as pptx_functions.show_functions() etc.
```

---

## Quick Start

```python
from pptx_functions import (
    create_slide_deck, get_default_config,
    OBJECT_TYPE_HANDLERS, color_to_rgb
)
from pptx.dml.color import RGBColor
from datetime import datetime, timezone

# ── 1. Define the slide deck ─────────────────────────────────────────────────
slide_config = {
    "Details": {
        "Author":             "K. Chadwick",
        "Title":              "Example Deck",
        "Filename":           "example.pptx",
        "Slide Aspect Ratio": "16:9",
        "Slide Width & Height": [13.33, 7.5],
    },
    "Slides": {
        "Slide 01": {
            "Slide Template":   6,          # 6 = blank slide
            "Slide Name":       "Cover",
            "Slide Notes":      "Auto-generated.",
            "Background Color": "#ffffff",
            "Background Alpha": 0.0,        # 0.0 = fully opaque
            "Objects": {
                "Title": get_default_config("Title", {
                    "Text": "Professor Hubert J. Farnsworth",
                }),
                "Image": get_default_config("Image", {
                    "img_path": "img/FuturamaProfessorFarnsworth.png",
                    "fit":      "height",
                    "left": 4.5, "top": 1.0, "height": 5.5,
                }),
            },
        }
    },
}

# ── 2. Create the presentation ───────────────────────────────────────────────
config_details = slide_config["Details"]
prs = create_slide_deck(config_details, verbose=True)

# ── 3. Build slides ──────────────────────────────────────────────────────────
for slide_name, config in slide_config["Slides"].items():
    layout = prs.slide_layouts[config.get("Slide Template", 6)]
    slide  = prs.slides.add_slide(layout)

    # Background colour
    fill = slide.background.fill
    fill.solid()
    r, g, b = color_to_rgb(config.get("Background Color", "#ffffff"))
    fill.fore_color.rgb          = RGBColor(r, g, b)
    fill.fore_color.transparency = config.get("Background Alpha", 0.0)

    # Add objects
    for key, elem_config in config["Objects"].items():
        if elem_config.get("Add?", False):
            obj_type = elem_config.get("Object Type", "")
            handler  = OBJECT_TYPE_HANDLERS.get(obj_type)
            if handler:
                handler(slide, elem_config)

# ── 4. Save ───────────────────────────────────────────────────────────────────
prs.save(config_details["Filename"])
```

---

## The `slide_config` Structure

```python
slide_config = {
    "Details": {                                # Presentation-level metadata
        "Author":               "K. Chadwick",
        "Title":                "Weekly Update",
        "Subject":              "Operations Summary",
        "Comments":             "Auto-generated by pptx_functions",
        "Keywords":             "pptx, automation",
        "Category":             "Workflow Automation",
        "Filename":             "Weekly Update 2026-04-29.pptx",
        "Slide Aspect Ratio":   "16:9",         # or "4:3"
        "Slide Width & Height": [13.33, 7.5],   # inches
    },

    "Slides": {
        "Slide 01": {
            "Slide Template":   6,              # 6 = blank (most common)
            "Slide Name":       "Cover Slide",
            "Slide Notes":      "Presenter notes go here.",
            "Background Color": "#1a1a2e",
            "Background Alpha": 0.0,

            "Objects": {
                "Title Config": { ... },        # any number of named objects
                "Image Config": { ... },
                "Line Config":  { ... },
            },
        },

        "Slide 02": { ... },
    },
}
```

Each entry in `"Objects"` is a config dict with an `"Object Type"` key that
tells the dispatch table which function to call.  Set `"Add?": False` to
temporarily skip an object without deleting its config.

### Slide Templates

| Index | Layout |
|---|---|
| `0` | Title Slide |
| `1` | Title and Content |
| `2` | Section Header |
| `3` | Two Content |
| `4` | Comparison |
| `5` | Title Only |
| `6` | **Blank** *(most common for programmatic use)* |
| `7` | Content with Caption |
| `8` | Picture with Caption |

---

## Object Configuration Examples

### Text Box — plain text

```python
"Title": {
    "Add?":        True,
    "Object Type": "Text",
    "Text":        "Good news, everyone!",
    "left": 1.5, "top": 0.3, "width": 10.0, "height": 0.6,
    "Align":       "left",
    "Font Name":   "Calibri",
    "Font Size":   28,
    "Font Color":  "#535353",
    "Bold?":       True,
    "Word Wrap?":  True,
    "Fill Color":  None,
    "Line Width":  0,
}
```

### Text Box — mixed-format runs

Different fonts, sizes, and colours within a single paragraph:

```python
"Status Line": {
    "Add?":        True,
    "Object Type": "Text",
    "Text": {
        "As of ":               {"Font Color": "#cb181d", "Bold?": True},
        "29 APR 26":            {"Font Color": "purple",  "Bold?": True, "Underline?": True, "Font Size": 16},
        ", the status is ":     {"Font Color": "#535353"},
        "OPERATIONAL":          {"Font Color": "green",   "Bold?": True, "Font Size": 18},
    },
    "left": 1.0, "top": 6.5, "width": 11.0, "height": 0.6,
    "Font Name": "Calibri", "Font Size": 12, "Font Color": "#535353",
}
```

### Text Box — bulleted list

`"Font Size"` can be a scalar or a `[[level, size], ...]` list for
per-indent sizing:

```python
"Bullets": {
    "Add?":        True,
    "Object Type": "Text",
    "Bullets": [
        ["Key Findings",                     0],
        ["Network traffic up 23% week-on-week", 1],
        ["Three new nodes identified in AO",    1],
        ["Pattern of Life",                  0],
        ["Activity concentrated 0200–0600Z",    1],
    ],
    "left": 0.5, "top": 1.2, "width": 6.0, "height": 5.5,
    "Font Name":  "Calibri",
    "Font Size":  [[0, 14], [1, 11]],       # level → size mapping
    "Font Color": "#252525",
    "Align":      "left",
}
```

### Image

```python
"Map": {
    "Add?":        True,
    "Object Type": "Image",
    "img_path":    "img/area_of_operations.png",
    "fit":         "width",         # "width" | "height" | "native"
    "Preserve Aspect Ratio?": True,
    "left": 6.5, "top": 1.0, "width": 6.5, "height": 5.5,
    "Line Width": 1,
    "Line Color": "#535353",
    "Line Style": "-",
}
```

### Connector

```python
"Divider": {
    "Add?":        True,
    "Object Type": "Connector",
    "Type":        "straight",      # "straight" | "elbow" | "curved"
    "Start X": 0.5, "Start Y": 1.0,
    "End X":   12.8, "End Y":  1.0,
    "Color":   "#535353",
    "Width":   1.5,
    "Style":   "-",                 # see show_dash_styles()
}
```

### Auto Shape

```python
"Highlight": {
    "Add?":          True,
    "Object Type":   "AutoShape",
    "AutoShape Key": "oval",        # see show_autoshapes()
    "left": 2.0, "top": 2.0, "width": 1.5, "height": 1.5,
    "Fill Color":    "#4169E1",
    "Fill Alpha":    0.6,
    "Line Color":    "#ffffff",
    "Line Width":    2,
}
```

### Table

```python
"Data Table": {
    "Add?":           True,
    "Object Type":    "Table",
    "left": 0.5, "top": 1.5, "width": 12.3, "height": 4.0,
    "Columns":        4,
    "Rows":           3,
    "Column Widths":  [3.0, 3.0, 3.0, 3.3],
    "Row Height":     0.4,
    "Column Headers": ["Location", "Date", "Activity", "Confidence"],
    "Row Data": [
        ["Grid 38SMB", "29 APR 26", "Vehicle movement", "High"],
        ["Grid 38SNC", "28 APR 26", "Signal intercept",  "Medium"],
        ["Grid 38SMD", "27 APR 26", "Personnel activity","Low"],
    ],
    "Font Size":  10,
    "Font Color": "#252525",
    "Align":      "center",
    "V-Align":    "middle",
}
```

### Banner

```python
"Classification": {
    "Add?":               True,
    "Object Type":        "Banner",
    "Text":               "UNCLASSIFIED",
    "Slide Aspect Ratio": "16:9",   # controls slide width
    "Font Color":         "#ffffff",
    "Font Size":          12,
    "Fill Color":         "#007a33",
}
```

---

## Using `get_default_config`

`get_default_config(object_type, overrides)` returns a deep copy of the
built-in defaults for any object type, with your overrides merged on top.
Use it to avoid writing full configs for simple objects:

```python
# Full default — just change what you need
title = get_default_config("Title", {"Text": "My Slide Title"})

# Default image at a custom position
img   = get_default_config("Image", {
    "img_path": "img/logo.png",
    "left": 0.2, "top": 0.2, "width": 1.5,
})

# Available object types
# "AutoShape", "Banner", "Connector", "Image", "Table", "Text", "Title"
```

---

## Discovery Helper Functions

```python
from pptx_functions import (
    show_functions,        # print all functions and data objects
    show_autoshapes,       # print valid "AutoShape Key" values
    show_object_alignment, # print valid "Align" values
    show_dash_styles,      # print valid "Line Style" values
    show_slide_templates,  # print slide layout descriptions
)

show_functions()       # full module quick-reference
show_autoshapes()      # oval, hexagon, star 7-point, ...
show_dash_styles()     # solid, dash, dash dot, round dot, ...
```

---

## Function Reference

| Function | Description |
|---|---|
| `create_slide_deck(config_details, verbose=False)` | Create a new Presentation with slide dimensions and metadata. |
| `add_autoshape(slide, config)` | Add an auto shape. |
| `add_connector(slide, config)` | Add a connector line. |
| `add_image(slide, config)` | Add an image with aspect-ratio handling. |
| `add_notes(slide, note_text)` | Write text to the slide Notes pane. |
| `add_shape_formatting(shape, config)` | Apply line and fill formatting to a shape. |
| `add_table(slide, config)` | Add a formatted table. |
| `add_textbox(slide, config)` | Add a text box (plain, multi-run, or bulleted). |
| `add_text_formatting(run, config)` | Apply font formatting to a text run. |
| `add_banners(slide, config)` | Add top and bottom marking banners. |
| `add_header(slide, config)` | Add a title + connector + seal image header. |
| `get_default_config(object_type, overrides)` | Return a deep-copied default config. |
| `apply_metadata(core_props, metadata_dict)` | Apply document metadata. |
| `color_to_rgb(color)` | Convert HEX or CSS name to (R, G, B) tuple. |
| `get_image_details(img_path)` | Extract image size, DPI, format via Pillow. |
| `extract_ltwh(config)` | Return [left, top, width, height] as Inches. |

---

## Dependencies

| Package | Purpose |
|---|---|
| [python-pptx](https://python-pptx.readthedocs.io/) | PowerPoint file generation |
| [Pillow](https://pillow.readthedocs.io/) | Image metadata extraction (size, DPI, format) |
| [webcolors](https://webcolors.readthedocs.io/) | CSS colour name → RGB conversion |

```bash
pip install python-pptx Pillow webcolors
```

---

## Example Output

The Professor Farnsworth slide generated by the Quick Start example:

<img src="img/FuturamaProfessorFarnsworth.png" width="480">

---

## References

- [python-pptx Documentation](https://python-pptx.readthedocs.io/en/latest/index.html)
- [python-pptx Shapes](https://python-pptx.readthedocs.io/en/latest/user/shapes.html)
- [MSO_SHAPE enum values](https://python-pptx.readthedocs.io/en/latest/api/enum/MsoAutoShapeType.html)
