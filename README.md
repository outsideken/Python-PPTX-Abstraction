# pptx_functions

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![python-pptx](https://img.shields.io/badge/python--pptx-0.6%2B-orange)
![License](https://img.shields.io/badge/license-MIT-green)

A Python library that separates **slide formatting** from **slide content** — so
you can define your visual style once and regenerate your deck as many times as
you need, with fresh data.

---

## The Core Idea

Every data-driven slide deck has two concerns:

**Style** — background colours, fonts, positions, column widths, border styles.
This is your house style. It changes rarely, maybe never.

**Content** — titles, body text, table rows, bullet points, image paths.
This is your data. It changes every time.

Most tools force you to write both in the same place.  `pptx_functions`
separates them.

You define your slide layout once as a Python dictionary — the **template**.
Then you fill in the content fields (plain strings, lists, and paths) and call
`build_presentation()`.  Every run produces a visually identical slide with
fresh data.

```
Template (formatting, fixed)  +  Content (data, changes)  →  .pptx
```

---

## Why a Dictionary?

The `slide_config` dictionary is the interface between your data and your
slides.  Every object on every slide — text boxes, tables, images, connectors —
is a plain dict with human-readable keys.

Because the keys are self-documenting (`"Font Color"`, `"Row Data"`,
`"Bullets"`), you can hand the content fields to **anything that can populate a
Python dictionary**: a script that queries a database, a function that parses a
CSV, or a large language model.

The consumer of those fields never needs to know what `RGBColor` is, what
`Inches(0.3)` means, or how python-pptx structures a table cell.  It just
returns strings and lists.

---

## Using with LLMs

This is where the design pays off most.  The pattern is:

1. **Define a template** — all formatting fixed, content fields left empty.
2. **Prompt an LLM** with a minimal JSON schema describing only the content
   fields.
3. **Inject the LLM's output** into a deep copy of the template.
4. **Call `build_presentation()`.**

The LLM never sees python-pptx.  It never touches a colour code or a pixel
coordinate.  It only fills in the fields you ask for.

```python
from copy import deepcopy
from pptx_functions import build_presentation, get_default_config

# ── Step 1: Define the template once ────────────────────────────────────────
# Formatting is fully specified. Content fields are intentionally empty.

template = {
    "Details": {
        "Author":             "B. Rodriguez",
        "Filename":           "weekly_summary.pptx",
        "Slide Aspect Ratio": "16:9",
        "Slide Width & Height": [13.33, 7.5],
    },
    "Slides": {
        "Slide 01": {
            "Slide Template":   6,
            "Background Color": "#1a1a2e",
            "Background Alpha": 0.0,
            "Objects": {
                "Banner": get_default_config("Banner", {
                    "Text":       "UNCLASSIFIED",
                    "Font Color": "#ffffff",
                    "Fill Color": "#007a33",
                }),
                "Title": get_default_config("Title", {
                    "Text":       "",           # ← content: LLM fills this
                    "Font Color": "#ffffff",
                }),
                "Summary": get_default_config("Text", {
                    "Text":      "",            # ← content: LLM fills this
                    "left": 0.5, "top": 1.2, "width": 6.0, "height": 4.5,
                    "Font Color": "#cccccc", "Font Size": 12, "Align": "left",
                }),
                "Activity Table": get_default_config("Table", {
                    "Column Headers": ["Location", "Date", "Activity", "Confidence"],
                    "Row Data":       [],        # ← content: LLM fills this
                    "Columns": 4, "Rows": 0,
                    "Column Widths": [2.5, 1.5, 4.5, 1.8],
                    "left": 6.8, "top": 1.2, "width": 6.0, "height": 4.5,
                    "Cell Styles": {
                        (0, 3): {"Font Color": "#ffffff", "Fill Color": "#1a1a2e"},
                    },
                }),
            },
        },
    },
}

# ── Step 2: Prompt an LLM for content ───────────────────────────────────────
# The prompt describes only the content schema.  Example:
#
#   Return a JSON object with exactly these fields, populated from the data
#   below.  Do not add or rename fields.
#
#   {
#     "title":   "<slide title including the reporting date>",
#     "summary": "<2–3 sentence narrative summary of key findings>",
#     "rows":    [["<location>", "<date>", "<activity>", "<High|Medium|Low>"], ...]
#   }
#
#   Data: [your source data here]

# LLM returns — no PowerPoint knowledge required:
llm_output = {
    "title":   "Weekly Activity Summary — 29 APR 2026",
    "summary": "Three locations reported elevated activity this week. "
               "Grid 38SMB shows continued vehicle movement consistent with "
               "prior pattern-of-life.  Two new nodes identified in AO.",
    "rows": [
        ["Grid 38SMB", "29 APR 26", "Vehicle movement",   "High"],
        ["Grid 38SNC", "28 APR 26", "Signal intercept",   "Medium"],
        ["Grid 38SMD", "27 APR 26", "Personnel activity", "Low"],
    ],
}

# ── Step 3: Inject content into a copy of the template ──────────────────────
slide_config = deepcopy(template)
objs = slide_config["Slides"]["Slide 01"]["Objects"]

objs["Title"]["Text"]                = llm_output["title"]
objs["Summary"]["Text"]              = llm_output["summary"]
objs["Activity Table"]["Row Data"]   = llm_output["rows"]
objs["Activity Table"]["Rows"]       = len(llm_output["rows"])

# ── Step 4: Build ────────────────────────────────────────────────────────────
prs = build_presentation(slide_config)
prs.save("weekly_summary.pptx")
```

The template lives in a JSON file or a version-controlled Python module.
The content comes from wherever your data lives.  **The formatting never
changes unless you change the template.**

Run it once a week, once an hour, or on every API response — the output is
always pixel-perfect to your spec.

---

## What It Supports

| Object Type | Key | Description |
|---|---|---|
| Text box | `"Text"` | Plain text, multi-paragraph, mixed-format runs, or bulleted lists |
| Image | `"Image"` | Smart aspect-ratio handling (`fit: width / height / native`) |
| Connector | `"Connector"` | Straight, elbow, or curved lines with dash styling |
| Auto Shape | `"AutoShape"` | 35+ MSO shapes with fill, transparency, and border formatting |
| Table | `"Table"` | Header row + data rows with optional per-cell styling |
| Banner | `"Banner"` | Top and bottom classification / marking banners |
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

The minimal end-to-end example:

```python
from pptx_functions import build_presentation, get_default_config

slide_config = {
    "Details": {
        "Author":             "B. Rodriguez",
        "Title":              "Example Deck",
        "Filename":           "example.pptx",
        "Slide Aspect Ratio": "16:9",
        "Slide Width & Height": [13.33, 7.5],
    },
    "Slides": {
        "Slide 01": {
            "Slide Template":   6,
            "Slide Name":       "Cover",
            "Slide Notes":      "Auto-generated.",
            "Background Color": "#ffffff",
            "Background Alpha": 0.0,
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

prs = build_presentation(slide_config, verbose=True)
prs.save(slide_config["Details"]["Filename"])
```

---

## The `slide_config` Structure

```python
slide_config = {
    "Details": {                                # Presentation-level metadata
        "Author":               "B. Rodriguez",
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

### Text Box — multi-paragraph

Each element in the list becomes its own paragraph.  Elements can be plain
strings or mixed-format run dicts:

```python
"Analysis": {
    "Add?":        True,
    "Object Type": "Text",
    "Text": [
        "Key Judgement:  Activity levels remain elevated.",
        {
            "Confidence: ": {"Bold?": True, "Font Color": "#535353"},
            "HIGH":         {"Bold?": True, "Font Color": "#007a33", "Font Size": 14},
        },
        "No change to recommended posture at this time.",
    ],
    "left": 0.5, "top": 1.2, "width": 6.0, "height": 3.0,
    "Font Name": "Calibri", "Font Size": 11, "Font Color": "#252525", "Align": "left",
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
        ["Key Findings",                        0],
        ["Network traffic up 23% week-on-week", 1],
        ["Three new nodes identified in AO",    1],
        ["Pattern of Life",                     0],
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
    "Line Color": "#535353",
    "Line Width": 1.5,
    "Line Style": "-",              # see show_dash_styles()
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
    "Cell Styles": {                 # optional per-cell overrides
        (0, 3): {"Font Color": "#ffffff", "Fill Color": "#1a1a2e"},
        (1, 3): {"Font Color": "#007a33", "Bold?": True},  # High → green
        (2, 3): {"Font Color": "#e6a817", "Bold?": True},  # Medium → amber
        (3, 3): {"Font Color": "#cb181d", "Bold?": True},  # Low → red
    },
}
```

### Banner

```python
"Classification": {
    "Add?":       True,
    "Object Type": "Banner",
    "Text":        "UNCLASSIFIED",
    "Font Color":  "#ffffff",
    "Font Size":   12,
    "Fill Color":  "#007a33",
}
```

---

## Using `get_default_config`

`get_default_config(object_type, overrides)` returns a deep copy of the
built-in defaults for any object type, with your overrides merged on top.
It is the recommended way to build template objects — write only the fields
that differ from the default:

```python
# Full default — just change what you need
title = get_default_config("Title", {"Text": "My Slide Title"})

# Default image at a custom position
img   = get_default_config("Image", {
    "img_path": "img/logo.png",
    "left": 0.2, "top": 0.2, "width": 1.5,
})

# Available object types:
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

**High-level API**

| Function | Description |
|---|---|
| `build_presentation(slide_config, verbose=False)` | Build a complete Presentation from a `slide_config` dict. |
| `dispatch_objects(slide, objects_config)` | Render all enabled objects from an `"Objects"` config dict. |

**PPTX Object Functions**

| Function | Description |
|---|---|
| `create_slide_deck(config_details, verbose=False)` | Create a new Presentation with slide dimensions and metadata. |
| `set_slide_background(slide, config)` | Apply background colour and transparency to a slide. |
| `add_autoshape(slide, config)` | Add an auto shape. |
| `add_connector(slide, config)` | Add a connector line. |
| `add_image(slide, config)` | Add an image with aspect-ratio handling. |
| `add_notes(slide, note_text)` | Write text to the slide Notes pane. |
| `add_shape_formatting(shape, config)` | Apply line and fill formatting to a shape. |
| `add_table(slide, config)` | Add a formatted table with optional per-cell styling. |
| `add_textbox(slide, config)` | Add a text box (plain, multi-paragraph, multi-run, or bulleted). |
| `add_text_formatting(run, config)` | Apply font formatting to a text run. |
| `add_banners(slide, config)` | Add top and bottom marking banners. |
| `add_header(slide, config)` | Add a title + connector + seal image header. |

**Helpers & Utilities**

| Function | Description |
|---|---|
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

## References

- [python-pptx Documentation](https://python-pptx.readthedocs.io/en/latest/index.html)
- [python-pptx Shapes](https://python-pptx.readthedocs.io/en/latest/user/shapes.html)
- [MSO_SHAPE enum values](https://python-pptx.readthedocs.io/en/latest/api/enum/MsoAutoShapeType.html)
