# pptx_functions

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![python-pptx](https://img.shields.io/badge/python--pptx-0.6%2B-orange)
![License](https://img.shields.io/badge/license-MIT-green)

A Python library for automating the production of data-heavy PowerPoint slides —
without removing the human judgment that makes those slides worth reading.

The design decisions, the visual style, the stakeholder approval — those stay
with your team.  The part that was always purely mechanical — opening the
template, pasting in this week's numbers, re-checking the formatting — that's
what this handles.

---

## The Workflow

```
Design in PowerPoint  →  extract_presentation_config()  →  template dict
                                                                  ↓
                                                    content fields = placeholders
                                                                  ↓
                                              inject data  ←  any source
                                          (database / CSV / LLM)
                                                                  ↓
                                                    build_presentation()
                                                                  ↓
                                                   pixel-perfect  .pptx
```

There are four steps.  A person is responsible for three of them.  The library
handles one.

---

### Step 1 — Design your slide in PowerPoint

Build the slide the way you want it.  Use the fonts, colors, and layout that
meet your organization's standards.  Get it approved.  This is entirely a human
step — the tool makes no aesthetic decisions.

---

### Step 2 — Extract the template

`extract_presentation_config()` reads your `.pptx` file and produces a Python
dictionary that captures every formatting detail: positions, fonts, colors,
column widths, line styles.  Image paths become `None` — a placeholder
signaling "content goes here."  Text becomes an empty string you'll replace.

```python
from pptx_functions import extract_presentation_config

template = extract_presentation_config("approved_briefing.pptx")
```

The result is a plain dict.  Save it as JSON, commit it to version control,
share it with your team.  The approved design is now locked — it only changes
if a person deliberately changes it.

```python
# What you get back — every formatting detail captured:
{
    "Details": {
        "Author":               "B. Rodriguez",
        "Slide Aspect Ratio":   "16:9",
        "Slide Width & Height": [13.33, 7.5],
        ...
    },
    "Slides": {
        "Slide 01": {
            "Background Color": "#0d1b2a",
            "Objects": {
                "Title": {
                    "Object Type": "Text",
                    "Text":        "",          # ← placeholder: inject your data here
                    "left": 1.41, "top": 0.42, "width": 10.5, "height": 0.5,
                    "Font Name":   "Calibri",
                    "Font Size":   28.0,
                    "Font Color":  "#f5c518",
                    ...
                },
                "Mission Log": {
                    "Object Type":    "Table",
                    "Column Headers": ["Destination", "Date", "Cargo Manifest", "Danger Level"],
                    "Row Data":       [],        # ← placeholder: inject your data here
                    ...
                },
                "Hero Image": {
                    "Object Type": "Image",
                    "img_path":    None,         # ← placeholder: inject your image path here
                    "left": 0.32, "top": 3.31, "width": 2.19, "height": 3.6,
                    ...
                },
            }
        }
    }
}
```

The extractor accepts either a file path or an open Presentation object:

```python
template = extract_presentation_config("approved_briefing.pptx")       # path string

from pptx import Presentation
template = extract_presentation_config(Presentation("approved_briefing.pptx"))  # or object
```

---

### Step 3 — Inject your content

The config separates formatting (everything extracted in Step 2 — fixed, never
changes) from content (the placeholder fields — replaced on every run).

Content can come from anywhere:

```python
from copy import deepcopy

slide_config = deepcopy(template)
objs = slide_config["Slides"]["Slide 01"]["Objects"]

# From a database:
objs["Title"]["Text"]           = db.get_title()
objs["Mission Log"]["Row Data"] = db.get_missions()
objs["Mission Log"]["Rows"]     = len(objs["Mission Log"]["Row Data"])

# From a file path:
objs["Hero Image"]["img_path"]  = "img/this_weeks_photo.png"
```

---

### Step 4 — Build

```python
from pptx_functions import build_presentation

prs = build_presentation(slide_config)
prs.save("briefing_2026-04-30.pptx")
```

Every run produces a visually identical slide with fresh data.  The formatting
never drifts.  The approved design stays approved.

---

## Using LLMs as a Data Source

Because the config uses plain English keys (`"Text"`, `"Row Data"`,
`"img_path"`), a large language model can fill in content fields without any
knowledge of PowerPoint, python-pptx, or slide coordinates.

You define the content schema — the LLM fills it in.  The LLM never touches a
colour code, a position, or a font name.  That's what makes it deterministic:
the formatting is fixed in the template, and the LLM's output goes only into
the fields you've explicitly opened up.

A person writes the prompt.  A person reviews the output.  The LLM handles the
text generation.  The library handles the formatting.  Each does what it's good
at.

```python
from copy import deepcopy
from pptx_functions import extract_presentation_config, build_presentation

# Load the approved template
template = extract_presentation_config("approved_briefing.pptx")

# ── Step 2: Prompt an LLM for content only ───────────────────────────────────
#
#   You are Professor Hubert J. Farnsworth briefing the Planet Express crew.
#   Return a JSON object with exactly these fields. Do not add or rename them.
#
#   {
#     "title":    "<briefing title including the stardate>",
#     "briefing": "<2-3 sentences starting with 'Good news, Everyone!'>",
#     "rows":     [["<destination>", "<stardate>", "<cargo>", "<High|Medium|Low>"], ...]
#   }
#
#   Mission data: [your source data here]

# LLM returns — no python-pptx knowledge required:
llm_output = {
    "title":    "Planet Express Weekly Mission Briefing — Stardate 29 APR 3000",
    "briefing": "Good news, Everyone!  I've assigned you three new deliveries to "
                "destinations that will, in all likelihood, result in at least one "
                "of your deaths.  The good news is I have no strong attachment to "
                "any of you.  Now stop touching my things.",
    "rows": [
        ["Omicron Persei 8",  "29 APR 3000", "1 can of anchovies (urgent)",     "High"],
        ["Robot Hell",        "30 APR 3000", "Bender's soul (retrieve, again)", "Medium"],
        ["Planet Nude Beach", "01 MAY 3000", "Swimsuit calendars — rush order", "Low"],
    ],
}

# ── Step 3: Inject into a copy of the template ───────────────────────────────
slide_config = deepcopy(template)
objs = slide_config["Slides"]["Slide 01"]["Objects"]

objs["Title"]["Text"]           = llm_output["title"]
objs["Bubble"]["Text"]          = llm_output["briefing"]
objs["Mission Log"]["Row Data"] = llm_output["rows"]
objs["Mission Log"]["Rows"]     = len(llm_output["rows"])

# Post-process: colour-code the Danger Level column from the LLM's values
danger_colors = {"High": "#cb181d", "Medium": "#e6a817", "Low": "#007a33"}
for i, row in enumerate(llm_output["rows"], start=1):
    objs["Mission Log"]["Cell Styles"][(i, 3)] = {
        "Font Color": danger_colors.get(row[3], "#535353"),
        "Bold?": True,
    }

# ── Step 4: Build ─────────────────────────────────────────────────────────────
prs = build_presentation(slide_config)
prs.save("mission_briefing_2026-04-30.pptx")
```

---

## What It Supports

| Object Type | `"Object Type"` key | Description |
|---|---|---|
| Text box | `"Text"` | Plain text, multi-paragraph, mixed-format runs, or bulleted lists |
| Image | `"Image"` | Smart aspect-ratio handling (`fit: width / height / native`), flip / mirror |
| Connector | `"Connector"` | Straight, elbow, or curved lines with dash styling |
| Auto Shape | `"AutoShape"` | 35+ MSO shapes including callouts; optional text rendered directly into the shape |
| Table | `"Table"` | Header row + data rows with full per-cell styling |
| Banner | `"Banner"` | Matching top and bottom text bars spanning the full slide width |
| Header | `"Header"` | Title text box + separator line + left/right seal images |

---

## Installation

```bash
pip install python-pptx Pillow webcolors
```

Copy `pptx_functions.py` into your project folder (or anywhere on your Python
path), then import it:

```python
from pptx_functions import *
```

---

## Quick Start

Build a slide from scratch without extracting an existing file:

```python
from pptx_functions import build_presentation, get_default_config

slide_config = {
    "Details": {
        "Author":               "B. Rodriguez",
        "Title":                "Example Deck",
        "Filename":             "example.pptx",
        "Slide Aspect Ratio":   "16:9",
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
                    "img_path": "img/ProfessorFarnsworth.png",
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
        "Filename":             "Weekly Update 2026-04-30.pptx",
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
                "Title": { ... },               # any number of named objects
                "Image": { ... },
                "Divider": { ... },
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
        "As of ":           {"Font Color": "#cb181d", "Bold?": True},
        "29 APR 26":        {"Font Color": "purple", "Bold?": True, "Underline?": True, "Font Size": 16},
        ", the status is ": {"Font Color": "#535353"},
        "OPERATIONAL":      {"Font Color": "green",  "Bold?": True, "Font Size": 18},
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
    "Font Size":  [[0, 14], [1, 11]],
    "Font Color": "#252525",
    "Align":      "left",
}
```

### Image

`"Flip Horizontal?"` and `"Flip Vertical?"` mirror the image without affecting
its position or bounding box:

```python
"Farnsworth": {
    "Add?":                   True,
    "Object Type":            "Image",
    "img_path":               "img/ProfessorFarnsworth.png",
    "Preserve Aspect Ratio?": True,
    "fit":                    "height",     # "width" | "height" | "native"
    "left": 0.32, "top": 3.31, "width": 2.19, "height": 3.6,
    "Flip Horizontal?":       True,
    "Flip Vertical?":         False,
    "Line Width": 0,
}
```

### Connector

```python
"Divider": {
    "Add?":        True,
    "Object Type": "Connector",
    "Type":        "straight",      # "straight" | "elbow" | "curved"
    "Start X": 0.5, "Start Y": 1.0,
    "End X":  12.8, "End Y":   1.0,
    "Line Color": "#535353",
    "Line Width": 1.5,
    "Line Style": "-",              # see show_dash_styles()
}
```

### Auto Shape

Shapes can be purely visual or include text rendered directly into the shape's
text frame — useful for callouts and speech bubbles:

```python
# Visual only
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

# With text — callout / speech bubble
"Bubble": {
    "Add?":          True,
    "Object Type":   "AutoShape",
    "AutoShape Key": "oval callout",
    "Text":          "Good news, Everyone!",
    "left": 2.0, "top": 2.88, "width": 5.5, "height": 1.3,
    "Align":         "center",
    "V Align":       "middle",
    "Font Name":     "Chalkboard",
    "Font Size":     10,
    "Font Color":    "#1a1a2e",
    "Word Wrap?":    True,
    "Fill Color":    "#b3d7f5",
    "Line Color":    "#5b7fae",
    "Line Width":    1.5,
}
```

Available callout shapes: `"oval callout"`, `"rounded callout"`,
`"rectangular callout"`, `"cloud callout"`.

### Table

All text parameters (`"Font Name"`, `"Font Size"`, `"Font Color"`, `"Bold?"`,
`"Italic?"`, `"Underline?"`, `"Align"`, `"V-Align"`) apply table-wide and can
be overridden per-cell via `"Cell Styles"`:

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
        ["Grid 38SMB", "29 APR 26", "Vehicle movement",  "High"],
        ["Grid 38SNC", "28 APR 26", "Signal intercept",   "Medium"],
        ["Grid 38SMD", "27 APR 26", "Personnel activity", "Low"],
    ],
    "Font Size":  10,
    "Font Color": "#252525",
    "Align":      "center",
    "V-Align":    "middle",
    "Cell Styles": {
        (0, 3): {"Font Color": "#ffffff", "Fill Color": "#1a1a2e"},
        (1, 3): {"Font Color": "#007a33", "Bold?": True},
        (2, 3): {"Font Color": "#e6a817", "Bold?": True},
        (3, 3): {"Font Color": "#cb181d", "Bold?": True},
    },
}
```

### Banner

Renders two identical text bars — one flush with the top edge of the slide,
one flush with the bottom — spanning the full slide width automatically.
Useful for classification markings or persistent labeling:

```python
"Banner": {
    "Add?":        True,
    "Object Type": "Banner",
    "Text":        "PLANET EXPRESS — FOR EXTERNAL USE ONLY",
    "Font Color":  "#ffffff",
    "Font Size":   14,
    "Bold?":       True,
    "Fill Color":  "#6a0572",
}
```

All standard text formatting keys are supported (`"Font Name"`, `"Font Color"`,
`"Font Size"`, `"Bold?"`, `"Italic?"`, `"Fill Color"`, `"Line Color"`, etc.).
`"width"` and both `"top"` values are set automatically from the slide
dimensions — there is no need to specify them.

---

## Using `get_default_config`

`get_default_config(object_type, overrides)` returns a deep copy of the
built-in defaults for any object type, with your overrides merged on top.
Write only the fields that differ from the default:

```python
title = get_default_config("Title", {"Text": "My Slide Title"})

img = get_default_config("Image", {
    "img_path": "img/logo.png",
    "left": 0.2, "top": 0.2, "width": 1.5,
})

# Available types: "AutoShape", "Connector", "Image", "Table", "Text", "Title"
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
```

---

## Function Reference

**High-level API**

| Function | Description |
|---|---|
| `build_presentation(slide_config, verbose=False)` | Build a complete Presentation from a `slide_config` dict. |
| `extract_presentation_config(prs)` | Extract a `slide_config` dict from an existing `.pptx` file or Presentation object. |
| `dispatch_objects(slide, objects_config)` | Render all enabled objects from an `"Objects"` config dict. |

**PPTX Object Functions**

| Function | Description |
|---|---|
| `create_slide_deck(config_details, verbose=False)` | Create a new Presentation with slide dimensions and metadata. |
| `set_slide_background(slide, config)` | Apply background colour and transparency to a slide. |
| `add_autoshape(slide, config)` | Add an auto shape with optional fill, border, and text. |
| `add_connector(slide, config)` | Add a connector line. |
| `add_image(slide, config)` | Add an image with aspect-ratio handling and optional flip. |
| `add_notes(slide, note_text)` | Write text to the slide Notes pane. |
| `add_shape_formatting(shape, config)` | Apply line and fill formatting to a shape. |
| `add_banners(slide, config)` | Add matching top and bottom banners spanning the full slide width. |
| `add_table(slide, config)` | Add a formatted table with optional per-cell styling. |
| `add_textbox(slide, config)` | Add a text box (plain, multi-paragraph, multi-run, or bulleted). |
| `add_text_formatting(run, config)` | Apply font formatting to a text run. |
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
