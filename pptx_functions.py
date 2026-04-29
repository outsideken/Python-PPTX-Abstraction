"""
pptx_functions  v0.1.0
=======================
python-pptx abstraction layer for programmatic PowerPoint slide generation.

Provides a configuration-dictionary workflow for building slides without
writing python-pptx boilerplate.  Each slide object (text box, image,
connector, auto shape, table, banner, header) is driven by a plain Python
dict, making slide configurations easy to generate, serialise, and reuse.

Quick start
-----------
>>> from pptx_functions import create_slide_deck, get_default_config, OBJECT_TYPE_HANDLERS
>>> from datetime import datetime, timezone
>>> slide_config = {
...     "Details": {
...         "Author": "K. Chadwick",
...         "Filename": "test.pptx",
...         "Slide Aspect Ratio": "16:9",
...         "Slide Width & Height": [13.33, 7.5],
...     },
...     "Slides": {
...         "Slide 01": {
...             "Slide Template": 6,
...             "Slide Name": "Test Slide",
...             "Slide Notes": "",
...             "Background Color": "#ffffff",
...             "Background Alpha": 0.0,
...             "Objects": {
...                 "Title": get_default_config("Title", {"Text": "Hello World"}),
...             },
...         }
...     },
... }

Functions
---------
Utilities
    get_image_details       Extract image metadata (size, DPI, format) using Pillow.
    color_to_rgb            Convert a CSS colour name or HEX string to (R, G, B).
    extract_ltwh            Return [left, top, width, height] as Inches from a config dict.

PPTX Object Functions
    create_slide_deck       Create a new Presentation with configured dimensions and metadata.
    add_autoshape           Add an auto shape to a slide.
    add_connector           Add a connector line to a slide.
    add_image               Add an image with smart aspect-ratio handling.
    add_notes               Write text to the Notes pane of a slide.
    add_shape_formatting    Apply line and fill formatting to any shape.
    add_table               Add a formatted table to a slide.
    add_textbox             Add a text box (plain, multi-run, or bulleted).
    add_text_formatting     Apply font formatting to a text run or paragraph.

Wrapper Functions
    add_banners             Add top and bottom classification/marking banners to a slide.
    add_header              Add a slide header (title text box, connector, and seal images).

Helper Functions
    get_default_config      Return a deep-copied default config for a named object type.
    apply_metadata          Apply document metadata to a Presentation's core properties.
    show_functions          Print all public functions and data objects in this module.
    show_autoshapes         Print available auto shape key strings.
    show_object_alignment   Print available text alignment key strings.
    show_dash_styles        Print available line dash style key strings.
    show_slide_templates    Print slide layout template descriptions.

Data
    PPTX_LOOKUP             Consolidated enum lookup dict (align, valign, dash_styles,
                            shapes, connectors).
    OBJECT_TYPE_HANDLERS    Dispatch table mapping "Object Type" strings to functions.
    default_configurations  Default config dicts for every supported object type.
"""

from __future__ import annotations

__version__ = "0.1.0"
__author__  = "KChadwick"

__all__ = [
    # utilities
    "get_image_details", "color_to_rgb", "extract_ltwh",
    # pptx object functions
    "create_slide_deck",
    "add_autoshape", "add_connector", "add_image", "add_notes",
    "add_shape_formatting", "add_table", "add_textbox", "add_text_formatting",
    # wrapper functions
    "add_banners", "add_header",
    # helper functions
    "get_default_config", "apply_metadata",
    "show_functions", "show_autoshapes", "show_object_alignment",
    "show_dash_styles", "show_slide_templates",
    # data
    "PPTX_LOOKUP", "OBJECT_TYPE_HANDLERS", "default_configurations",
]

import os
from copy import deepcopy
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image
import pptx
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.parts.coreprops import CorePropertiesPart as CoreProperties
from pptx.shapes.autoshape import Shape
from pptx.shapes.connector import Connector
from pptx.shapes.picture import Picture
from pptx.slide import Slide
from pptx.table import Table, _Cell
from pptx.util import Inches, Pt
import webcolors


################################################################################
# LOOKUP DATA
# Consolidated enum mappings used by all PPTX object functions.
################################################################################

PPTX_LOOKUP: Dict[str, Dict[str, Any]] = {

    "align": {
        "center":      PP_ALIGN.CENTER,
        "left":        PP_ALIGN.LEFT,
        "right":       PP_ALIGN.RIGHT,
        "justify":     PP_ALIGN.JUSTIFY,
        "justify_low": PP_ALIGN.JUSTIFY_LOW,
        "mixed":       PP_ALIGN.MIXED,
    },

    "valign": {
        "top":    MSO_VERTICAL_ANCHOR.TOP,
        "middle": MSO_VERTICAL_ANCHOR.MIDDLE,
        "bottom": MSO_VERTICAL_ANCHOR.BOTTOM,
    },

    "dash_styles": {
        "solid":         MSO_LINE_DASH_STYLE.SOLID,
        "-":             MSO_LINE_DASH_STYLE.SOLID,
        "dash":          MSO_LINE_DASH_STYLE.DASH,
        "--":            MSO_LINE_DASH_STYLE.DASH,
        "dash dot":      MSO_LINE_DASH_STYLE.DASH_DOT,
        "-.":            MSO_LINE_DASH_STYLE.DASH_DOT,
        "dash dot dot":  MSO_LINE_DASH_STYLE.DASH_DOT_DOT,
        "-..":           MSO_LINE_DASH_STYLE.DASH_DOT_DOT,
        "long dash":     MSO_LINE_DASH_STYLE.LONG_DASH,
        "long dash dot": MSO_LINE_DASH_STYLE.LONG_DASH_DOT,
        "round dot":     MSO_LINE_DASH_STYLE.ROUND_DOT,
        ".":             MSO_LINE_DASH_STYLE.ROUND_DOT,
        "square dot":    MSO_LINE_DASH_STYLE.SQUARE_DOT,
    },

    "shapes": {
        "trapezoid":         MSO_SHAPE.TRAPEZOID,
        "cube":              MSO_SHAPE.CUBE,
        "parallelogram":     MSO_SHAPE.PARALLELOGRAM,
        "sun":               MSO_SHAPE.SUN,
        "moon":              MSO_SHAPE.MOON,
        "lightning bolt":    MSO_SHAPE.LIGHTNING_BOLT,
        "hexagon":           MSO_SHAPE.HEXAGON,
        "heptagon":          MSO_SHAPE.HEPTAGON,
        "octagon":           MSO_SHAPE.OCTAGON,
        "cloud":             MSO_SHAPE.CLOUD,
        "gear 6":            MSO_SHAPE.GEAR_6,
        "gear 9":            MSO_SHAPE.GEAR_9,
        "explosion 1":       MSO_SHAPE.EXPLOSION1,
        "explosion 2":       MSO_SHAPE.EXPLOSION2,
        "smiley face":       MSO_SHAPE.SMILEY_FACE,
        "heart":             MSO_SHAPE.HEART,
        "rectangle":         MSO_SHAPE.RECTANGLE,
        "rounded rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval":              MSO_SHAPE.OVAL,
        "diamond":           MSO_SHAPE.DIAMOND,
        "right triangle":    MSO_SHAPE.RIGHT_TRIANGLE,
        "bent arrow":        MSO_SHAPE.BENT_ARROW,
        "right arrow":       MSO_SHAPE.RIGHT_ARROW,
        "left arrow":        MSO_SHAPE.LEFT_ARROW,
        "up arrow":          MSO_SHAPE.UP_ARROW,
        "down arrow":        MSO_SHAPE.DOWN_ARROW,
        "left right arrow":  MSO_SHAPE.LEFT_RIGHT_ARROW,
        "up down arrow":     MSO_SHAPE.UP_DOWN_ARROW,
        "star 4-point":      MSO_SHAPE.STAR_4_POINT,
        "star 5-point":      MSO_SHAPE.STAR_5_POINT,
        "star 6-point":      MSO_SHAPE.STAR_6_POINT,
        "star 7-point":      MSO_SHAPE.STAR_7_POINT,
        "star 8-point":      MSO_SHAPE.STAR_8_POINT,
        "star 16-point":     MSO_SHAPE.STAR_16_POINT,
        "star 32-point":     MSO_SHAPE.STAR_32_POINT,
        "chevron":           MSO_SHAPE.CHEVRON,
        "pentagon":          MSO_SHAPE.PENTAGON,
        "default":           MSO_SHAPE.RECTANGLE,
    },

    "connectors": {
        "straight": MSO_CONNECTOR.STRAIGHT,
        "elbow":    MSO_CONNECTOR.ELBOW,
        "curved":   MSO_CONNECTOR.CURVE,
    },
}

_SLIDE_TEMPLATES: Dict[str, str] = {
    "0": "Title Slide — title and subtitle placeholders.",
    "1": "Title and Content — title at top, one content placeholder.",
    "2": "Section Header — single title placeholder.",
    "3": "Two Content — title at top, two side-by-side content placeholders.",
    "4": "Comparison — title, two content placeholders with headings.",
    "5": "Title Only — title placeholder, no content area.",
    "6": "Blank — no placeholders (most common for programmatic use).",
    "7": "Content with Caption — content, title, and caption placeholder.",
    "8": "Picture with Caption — picture, title, and caption placeholder.",
}


################################################################################
# UTILITIES
################################################################################

def get_image_details(img_path: str) -> Dict[str, Any]:
    """
    Extract metadata from an image file using Pillow.

    Parameters
    ----------
    img_path : str
        Path to the image file.  Supports all formats Pillow can read
        (JPEG, PNG, GIF, TIFF, WebP, BMP, ICO, etc.).

    Returns
    -------
    dict
        Keys: ``file_path``, ``file_format``, ``color_mode``, ``width_px``,
        ``height_px``, ``aspect_ratio``, ``size_pixels``, ``has_alpha``,
        ``is_animated``, ``n_frames``, ``metadata``, ``dpi``.

    Raises
    ------
    FileNotFoundError
        If *img_path* does not exist.
    """
    if not os.path.exists(img_path):
        raise FileNotFoundError(f"Image not found: {img_path!r}")

    with Image.open(img_path) as im:
        im.load()
        return {
            "file_path":    os.path.abspath(img_path),
            "file_format":  im.format,
            "color_mode":   im.mode,
            "width_px":     im.width,
            "height_px":    im.height,
            "aspect_ratio": im.width / im.height,
            "size_pixels":  im.size,
            "has_alpha":    im.mode in ("RGBA", "LA", "PA"),
            "is_animated":  getattr(im, "is_animated", False),
            "n_frames":     getattr(im, "n_frames", 1),
            "metadata":     dict(im.info),
            "dpi":          im.info.get("dpi", (72.0, 72.0)),
        }


def color_to_rgb(color: str) -> Tuple[int, int, int]:
    """
    Convert a CSS colour name or HEX code to an (R, G, B) integer tuple.

    Parameters
    ----------
    color : str
        A HEX string (e.g. ``"#cb181d"``) or CSS3 colour name (e.g. ``"red"``).

    Returns
    -------
    tuple of int
        ``(R, G, B)`` with each component in [0, 255].
        Returns ``(0, 0, 0)`` (black) and prints a warning for invalid input.

    Examples
    --------
    >>> color_to_rgb("#cb181d")
    (203, 24, 29)
    >>> color_to_rgb("steelblue")
    (70, 130, 180)
    """
    try:
        if color.startswith("#"):
            return webcolors.hex_to_rgb(color)
        return webcolors.name_to_rgb(color)
    except (ValueError, AttributeError):
        print(f"Warning: {color!r} is not a valid colour name or HEX code. Using black.")
        return (0, 0, 0)


def extract_ltwh(config: Dict[str, Any]) -> List:
    """
    Return ``[left, top, width, height]`` as ``Inches`` objects from *config*.

    Only keys present in *config* are included, so partial configs (e.g. a
    connector with only position keys) are handled correctly.

    Parameters
    ----------
    config : dict
        Must contain at minimum ``"left"`` and ``"top"``; ``"width"`` and
        ``"height"`` are included when present.

    Returns
    -------
    list of pptx.util.Inches
    """
    return [Inches(config[k]) for k in ("left", "top", "width", "height") if k in config]


################################################################################
# PPTX OBJECT FUNCTIONS
################################################################################

def create_slide_deck(
    config_details: Dict[str, Any],
    verbose: bool = False,
) -> Presentation:
    """
    Create a new PowerPoint presentation with configured dimensions and metadata.

    Sets the slide size, stamps the current UTC time as the creation date, and
    applies all recognised metadata fields.  Does not add any slides.

    Parameters
    ----------
    config_details : dict
        Presentation-level configuration.  Required keys:

        ``"Slide Width & Height"``
            ``[width_inches, height_inches]`` — use ``[13.33, 7.5]`` for 16:9
            or ``[10.0, 7.5]`` for 4:3.

        Optional keys (applied as core properties):
        ``"Author"``, ``"Title"``, ``"Subject"``, ``"Keywords"``,
        ``"Comments"``, ``"Category"``, ``"Description"``.

    verbose : bool, optional
        Print configuration details to stdout.  Default ``False``.

    Returns
    -------
    pptx.presentation.Presentation

    Examples
    --------
    >>> details = {"Slide Width & Height": [13.33, 7.5], "Author": "K. Chadwick"}
    >>> prs = create_slide_deck(details)
    """
    prs = Presentation()
    config_details["Created"] = datetime.now(timezone.utc)

    slide_width, slide_height = config_details["Slide Width & Height"]
    prs.slide_width  = Inches(slide_width)
    prs.slide_height = Inches(slide_height)

    if verbose:
        print("Configuration Details:")
        for key, val in config_details.items():
            print(f"  {key}: {val}")
        print(f"\n  Slide Width:  {slide_width:.4f} inches")
        print(f"  Slide Height: {slide_height:.4f} inches")
        print(f"  Aspect Ratio: {slide_width / slide_height:.3f}:1\n")

    apply_metadata(prs.core_properties, config_details)
    return prs


def add_autoshape(slide: Slide, config: Dict[str, Any]) -> Shape:
    """
    Add an auto shape to a slide.

    Parameters
    ----------
    slide : Slide
        Target slide.
    config : dict
        Configuration keys:

        ``"AutoShape Key"`` (str)
            Shape name — see :func:`show_autoshapes` for valid values.
            Defaults to ``"rectangle"``.
        ``"left"``, ``"top"``, ``"width"``, ``"height"`` (float)
            Position and size in inches.
        ``"Fill Color"`` (str, optional)
            HEX or CSS name for the shape fill.
        ``"Fill Alpha"`` (float, optional)
            Fill opacity 0.0–1.0.  Default ``1.0``.
        ``"Line Color"`` (str, optional)
            HEX or CSS name for the border.
        ``"Line Width"`` (float, optional)
            Border width in points.  Default ``0`` (no border).
        ``"Line Style"`` (str, optional)
            Dash style — see :func:`show_dash_styles`.  Default ``"-"`` (solid).

    Returns
    -------
    Shape
    """
    left, top, width, height = extract_ltwh(config)
    shape_type = PPTX_LOOKUP["shapes"].get(
        config.get("AutoShape Key", "default"), MSO_SHAPE.RECTANGLE
    )
    shape: Shape = slide.shapes.add_shape(
        autoshape_type_id=shape_type,
        left=left, top=top, width=width, height=height,
    )
    add_shape_formatting(shape, config)
    return shape


def add_connector(slide: Slide, config: Dict[str, Any]) -> Connector:
    """
    Add a connector line to a slide.

    Parameters
    ----------
    slide : Slide
        Target slide.
    config : dict
        Configuration keys:

        ``"Color"`` (str)
            Line colour as HEX or CSS name.  Default ``"#000000"``.
        ``"Start X"``, ``"Start Y"``, ``"End X"``, ``"End Y"`` (float)
            Endpoint coordinates in inches.
        ``"Width"`` (float, optional)
            Line thickness in points.  Default ``1``.
        ``"Style"`` (str, optional)
            Dash style.  Default ``"-"`` (solid).
        ``"Type"`` (str, optional)
            ``"straight"``, ``"elbow"``, or ``"curved"``.  Default ``"straight"``.

    Returns
    -------
    Connector
    """
    rgb      = color_to_rgb(config.get("Color", "#000000"))
    start_x  = Inches(config.get("Start X", 0))
    start_y  = Inches(config.get("Start Y", 0))
    end_x    = Inches(config.get("End X", 0))
    end_y    = Inches(config.get("End Y", 0))
    conn_type = PPTX_LOOKUP["connectors"].get(
        config.get("Type", "straight"), MSO_CONNECTOR.STRAIGHT
    )

    line: Connector = slide.shapes.add_connector(conn_type, start_x, start_y, end_x, end_y)
    line.line.color.rgb  = RGBColor(*rgb)
    line.line.width      = Pt(config.get("Width", 1))
    line.line.dash_style = PPTX_LOOKUP["dash_styles"].get(
        config.get("Style", "-"), MSO_LINE_DASH_STYLE.SOLID
    )
    return line


def add_image(slide: Slide, config: Dict[str, Any]) -> Picture:
    """
    Add an image to a slide with smart aspect-ratio handling.

    Aspect ratio behaviour is controlled by ``"Preserve Aspect Ratio?"`` and
    ``"fit"``:

    ========================  ==================================================
    ``"fit"`` value           Behaviour
    ========================  ==================================================
    ``"width"``               Scale to the given width; compute height.
    ``"height"``              Scale to the given height; compute width.
    ``"native"`` (default)    Use the image's native pixel size at 96 DPI.
    ========================  ==================================================

    When ``"Preserve Aspect Ratio?"`` is ``False``, both ``"width"`` and
    ``"height"`` are applied exactly (stretching if needed).

    Parameters
    ----------
    slide : Slide
        Target slide.
    config : dict
        Required keys: ``"img_path"``, ``"left"``, ``"top"``.

        Optional keys: ``"width"``, ``"height"``, ``"fit"``,
        ``"Preserve Aspect Ratio?"``, ``"Line Width"``, ``"Line Color"``,
        ``"Line Style"``.

    Returns
    -------
    Picture

    Raises
    ------
    FileNotFoundError
        If ``"img_path"`` does not point to an existing file.
    KeyError
        If any required key is missing from *config*.
    """
    for key in ("img_path", "left", "top"):
        if key not in config:
            raise KeyError(f"add_image: missing required config key '{key}'")

    left     = Inches(config["left"])
    top      = Inches(config["top"])
    details  = get_image_details(config["img_path"])
    aspect   = details["aspect_ratio"]
    native_w = details["width_px"] / 96.0
    native_h = details["height_px"] / 96.0

    preserve = config.get("Preserve Aspect Ratio?", True)
    fit_mode = config.get("fit", "native").lower()

    if preserve:
        if "width" in config and fit_mode in ("width", "native"):
            width  = Inches(config["width"])
            height = Inches(config["width"] / aspect)
        elif "height" in config and fit_mode == "height":
            height = Inches(config["height"])
            width  = Inches(config["height"] * aspect)
        else:
            width  = Inches(native_w)
            height = Inches(native_h)
    else:
        width  = Inches(config.get("width", native_w))
        height = Inches(config.get("height", native_h))

    picture: Picture = slide.shapes.add_picture(
        config["img_path"], left, top, width=width, height=height
    )

    # Support both Title Case ("Line Width") and legacy snake_case ("line_width")
    lw = config.get("Line Width", config.get("line_width", 0))
    if lw > 0:
        lc = config.get("Line Color", config.get("line_color", "#000000"))
        ls = config.get("Line Style", config.get("line_style", "-"))
        picture.line.color.rgb  = RGBColor(*color_to_rgb(lc))
        picture.line.width      = Pt(lw)
        picture.line.dash_style = PPTX_LOOKUP["dash_styles"].get(ls, MSO_LINE_DASH_STYLE.SOLID)

    return picture


def add_notes(slide: Slide, note_text: str) -> None:
    """
    Write text to the Notes pane of a slide.

    Parameters
    ----------
    slide : Slide
        Target slide.
    note_text : str
        Text to place in the notes pane.  Replaces any existing content.
    """
    slide.notes_slide.notes_text_frame.text = note_text


def add_shape_formatting(shape: Shape, config: Dict[str, Any]) -> Shape:
    """
    Apply line and fill formatting to a PowerPoint shape.

    Transparency is applied via an overlay rectangle workaround because
    python-pptx's ``fore_color.transparency`` is unreliable across
    PowerPoint versions.

    Parameters
    ----------
    shape : Shape
        The shape to format (AutoShape, TextBox, etc.).  Modified in place.
    config : dict
        Formatting keys:

        ``"Line Color"`` (str, optional)   Border colour.
        ``"Line Width"`` (float, optional) Border width in points; ``0`` = none.
        ``"Line Style"`` (str, optional)   Dash style.  Default ``"-"`` (solid).
        ``"Fill Color"`` (str, optional)   Fill colour.  ``None`` = transparent.
        ``"Fill Alpha"`` (float, optional) Opacity 0.0–1.0.  Default ``1.0``.

    Returns
    -------
    Shape
        The same shape object (modified in place).
    """
    # ── Line formatting ───────────────────────────────────────────────────────
    line_color = config.get("Line Color")
    line_width = config.get("Line Width", 0)
    if line_color and line_width > 0:
        shape.line.color.rgb  = RGBColor(*color_to_rgb(line_color))
        shape.line.width      = Pt(line_width)
        shape.line.dash_style = PPTX_LOOKUP["dash_styles"].get(
            config.get("Line Style", "-"), MSO_LINE_DASH_STYLE.SOLID
        )

    # ── Fill formatting ───────────────────────────────────────────────────────
    fill_color = config.get("Fill Color")
    if not fill_color:
        return shape

    r, g, b = color_to_rgb(fill_color)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(r, g, b)

    # Transparency via overlay rectangle (most reliable cross-version approach)
    alpha = config.get("Fill Alpha", 1.0)
    if 0.0 <= alpha < 0.999:
        try:
            slide = shape.part.slide
            overlay: Shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, shape.left, shape.top, shape.width, shape.height
            )
            overlay.fill.solid()
            overlay.fill.fore_color.rgb          = RGBColor(r, g, b)
            overlay.fill.fore_color.transparency = 1.0 - alpha
            overlay.line.fill.background()
        except Exception as exc:
            print(f"Warning: transparency overlay failed — {exc}")

    return shape


def add_table(slide: Slide, config: Dict[str, Any]) -> Table:
    """
    Add a formatted table to a slide.

    Parameters
    ----------
    slide : Slide
        Target slide.
    config : dict
        Configuration keys:

        ``"left"``, ``"top"``, ``"width"``, ``"height"`` (float)
            Overall table position and size in inches.
        ``"Rows"`` (int)
            Number of data rows (header row is added automatically).
        ``"Columns"`` (int)
            Number of columns.
        ``"Column Widths"`` (list of float)
            Width of each column in inches.
        ``"Row Height"`` (float)
            Height of each row in inches.  Default ``0.4``.
        ``"Column Headers"`` (list of str)
            Labels for the top row.
        ``"Row Data"`` (list of list)
            2-D list of cell values (str).
        ``"Align"`` (str, optional)
            Horizontal text alignment.  Default ``"left"``.
        ``"V-Align"`` (str, optional)
            Vertical cell alignment.  Default ``"middle"``.
        ``"Font Size"`` (int, optional)
            Font size in points.  Default ``10``.
        ``"Font Color"`` (str, optional)
            HEX or CSS colour for cell text.
        ``"Bold?"`` (bool, optional)
            ``True`` bolds all cells; header row is bold by default.

    Returns
    -------
    Table
    """
    left, top, width, height = extract_ltwh(config)
    rows_count = config.get("Rows", 0) + 1
    cols_count = config.get("Columns", 0)

    shape       = slide.shapes.add_table(rows_count, cols_count, left, top, width, height)
    table: Table = shape.table

    for idx, col_w in enumerate(config.get("Column Widths", [])):
        if idx < len(table.columns):
            table.columns[idx].width = Inches(col_w)

    for row in table.rows:
        row.height = Inches(config.get("Row Height", 0.4))

    align_key  = config.get("Align", "left").lower()
    valign_key = config.get("V-Align", "middle").lower()
    full_data  = [config.get("Column Headers", [])] + config.get("Row Data", [])

    for row_idx, row_values in enumerate(full_data):
        for col_idx, cell_value in enumerate(row_values):
            if col_idx >= cols_count:
                break
            cell: _Cell = table.cell(row_idx, col_idx)
            cell.vertical_anchor = PPTX_LOOKUP["valign"].get(
                valign_key, MSO_VERTICAL_ANCHOR.MIDDLE
            )
            tf = cell.text_frame
            tf.clear()
            p = tf.add_paragraph()
            p.alignment = PPTX_LOOKUP["align"].get(align_key, PP_ALIGN.LEFT)
            run = p.add_run()
            run.text = str(cell_value)
            is_header = row_idx == 0
            add_text_formatting(run, {**config, "Bold?": config.get("Bold?", is_header)})

    return table


def add_textbox(slide: Slide, config: Dict[str, Any]) -> Shape:
    """
    Add a text box to a slide.

    Supports three text modes selected automatically from the config:

    1. **Bulleted list** — when ``"Bullets"`` is present.
       Value: ``[[text, level], ...]`` where level 0 = top level.
       ``"Font Size"`` may be a scalar or a ``[[level, size], ...]`` list.

    2. **Multi-run text** — when ``"Text"`` is a ``dict``.
       Value: ``{run_text: {per-run overrides}, ...}``.
       Useful for mixed formatting within one paragraph (e.g. bold + colour).

    3. **Plain text** — when ``"Text"`` is a ``str`` (default).

    Parameters
    ----------
    slide : Slide
        Target slide.
    config : dict
        Configuration keys:

        ``"left"``, ``"top"``, ``"width"``, ``"height"`` (float)
            Position and size in inches.
        ``"Text"`` (str or dict)
            Text content.  Required unless ``"Bullets"`` is provided.
        ``"Bullets"`` (list of [str, int], optional)
            Overrides ``"Text"``; drives bulleted-list mode.
        ``"Align"`` (str, optional)
            Horizontal alignment.  Default ``"center"``.
        ``"V Align"`` (str, optional)
            Vertical alignment.  Default ``"middle"``.
        ``"Font Name"`` (str, optional)    Default ``"Calibri"``.
        ``"Font Size"`` (int or list)      Default ``18``.
        ``"Font Color"`` (str, optional)   Default ``"#535353"``.
        ``"Bold?"``, ``"Italic?"``, ``"Underline?"`` (bool, optional)
        ``"Word Wrap?"`` (bool, optional)  Default ``True``.
        ``"Fill Color"``, ``"Fill Alpha"`` — passed to :func:`add_shape_formatting`.
        ``"Line Color"``, ``"Line Width"``, ``"Line Style"`` — border options.

    Returns
    -------
    Shape
        The text box shape object.
    """
    left, top, width, height = extract_ltwh(config)
    tx_box: Shape = slide.shapes.add_textbox(left, top, width, height)
    add_shape_formatting(tx_box, config)

    tf = tx_box.text_frame
    tf.word_wrap      = config.get("Word Wrap?", True)
    tf.auto_size      = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = PPTX_LOOKUP["valign"].get(
        config.get("V Align", "middle").lower(), MSO_VERTICAL_ANCHOR.MIDDLE
    )
    tf.clear()

    align_key = config.get("Align", "center").lower()

    # ── Bulleted list ─────────────────────────────────────────────────────────
    if "Bullets" in config:
        font_size_raw = config.get("Font Size", 12)
        if isinstance(font_size_raw, (list, tuple)):
            font_size_map     = dict(font_size_raw)
            default_font_size = 12
        elif isinstance(font_size_raw, dict):
            font_size_map     = font_size_raw
            default_font_size = 12
        else:
            font_size_map     = {}
            default_font_size = font_size_raw

        for bullet_text, level in config["Bullets"]:
            p          = tf.add_paragraph()
            p.level    = level
            p.alignment = PPTX_LOOKUP["align"].get(align_key, PP_ALIGN.LEFT)
            run        = p.add_run()
            run.text   = bullet_text
            font_size  = font_size_map.get(level, default_font_size)
            add_text_formatting(run, {**config, "Font Size": font_size})

    # ── Multi-run text ────────────────────────────────────────────────────────
    elif isinstance(config.get("Text"), dict):
        p           = tf.add_paragraph()
        p.alignment = PPTX_LOOKUP["align"].get(align_key, PP_ALIGN.CENTER)
        for text, run_config in config["Text"].items():
            run      = p.add_run()
            run.text = text
            add_text_formatting(run, {**config, **run_config})

    # ── Plain text ────────────────────────────────────────────────────────────
    else:
        p           = tf.add_paragraph()
        p.alignment = PPTX_LOOKUP["align"].get(align_key, PP_ALIGN.CENTER)
        run         = p.add_run()
        run.text    = str(config.get("Text", ""))
        add_text_formatting(run, config)

    return tx_box


def add_text_formatting(run, config: Dict[str, Any]) -> None:
    """
    Apply font formatting to a text run or paragraph.

    Parameters
    ----------
    run : pptx run or paragraph
        The object whose ``.font`` attribute will be set.
    config : dict
        Formatting keys:

        ``"Font Name"`` (str)    Font family.  Default ``"Calibri"``.
        ``"Font Size"`` (float)  Size in points.  Default ``12``.
        ``"Font Color"`` (str)   HEX or CSS colour.  Default ``"#000000"``.
        ``"Bold?"`` (bool)       Default ``False``.
        ``"Italic?"`` (bool)     Default ``False``.
        ``"Underline?"`` (bool)  Default ``False``.
    """
    run.font.name      = config.get("Font Name", "Calibri")
    run.font.size      = Pt(config.get("Font Size", 12))
    run.font.bold      = config.get("Bold?", False)
    run.font.italic    = config.get("Italic?", False)
    run.font.underline = config.get("Underline?", False)

    font_color = config.get("Font Color", "#000000")
    if font_color:
        run.font.color.rgb = RGBColor(*color_to_rgb(font_color))


################################################################################
# WRAPPER FUNCTIONS
################################################################################

def add_banners(slide: Slide, config: Dict[str, Any]) -> None:
    """
    Add top and bottom classification/marking banners to a slide.

    Places two identical text boxes — one flush with the top edge, one flush
    with the bottom — spanning the full slide width.  Banner height and all
    text formatting are drawn from ``default_configurations["Banner"]`` and
    can be overridden via *config*.

    Parameters
    ----------
    slide : Slide
        Target slide.
    config : dict
        Keys:

        ``"Text"`` (str)
            Banner text (e.g. ``"UNCLASSIFIED"``).
        ``"Slide Aspect Ratio"`` (str, optional)
            ``"16:9"`` or ``"4:3"``.  Controls slide width.  Default ``"4:3"``.

        All other Text config keys (``"Font Name"``, ``"Font Color"``, etc.)
        are merged on top of the Banner defaults.
    """
    slide_width  = 13.33 if config.get("Slide Aspect Ratio", "4:3") == "16:9" else 10.0
    banner       = deepcopy(default_configurations["Banner"])
    banner.update({k: v for k, v in config.items() if k != "Add?"})
    banner["width"] = slide_width

    slide_height = 7.5
    for top in (0.0, slide_height - banner["height"]):
        cfg       = deepcopy(banner)
        cfg["top"] = top
        add_textbox(slide, cfg)


def add_header(slide: Slide, config: Dict[str, Any]) -> None:
    """
    Add a slide header consisting of a title text box, a separator connector,
    and optional left and right seal images.

    Parameters
    ----------
    slide : Slide
        Target slide.
    config : dict
        Keys:

        ``"Text"`` (str)
            Title text for the header.
        ``"Slide Aspect Ratio"`` (str, optional)
            ``"16:9"`` or ``"4:3"``.  Default ``"4:3"``.
        ``"Header Connector"`` (dict, optional)
            Connector config dict.  ``"End X"`` is computed automatically
            from the slide width and ``"Start X"``.
        ``"Left Seal"`` (dict, optional)
            Image config dict for the left seal/logo.
        ``"Right Seal"`` (dict, optional)
            Image config dict for the right seal/logo.
        ``"Add Right Image?"`` (bool, optional)
            Set ``False`` to suppress the right seal.  Default ``True``.
    """
    slide_width = 13.33 if config.get("Slide Aspect Ratio", "4:3") == "16:9" else 10.0

    title_cfg         = deepcopy(default_configurations["Title"])
    title_cfg["Text"] = config.get("Text", "")
    title_cfg["width"] = slide_width - 1.5
    add_textbox(slide, title_cfg)

    if "Header Connector" in config:
        conn_cfg        = deepcopy(config["Header Connector"])
        conn_cfg["End X"] = slide_width - conn_cfg.get("Start X", 0)
        add_connector(slide, conn_cfg)

    if "Left Seal" in config:
        add_image(slide, config["Left Seal"])

    if config.get("Add Right Image?", True) and "Right Seal" in config:
        add_image(slide, config["Right Seal"])


################################################################################
# HELPER FUNCTIONS
################################################################################

def get_default_config(
    object_type: str,
    overrides: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Return a deep-copied default configuration for a named object type.

    Parameters
    ----------
    object_type : str
        One of: ``"AutoShape"``, ``"Banner"``, ``"Connector"``, ``"Image"``,
        ``"Table"``, ``"Text"``, ``"Title"``.
    overrides : dict, optional
        Keys to update on top of the defaults.  The defaults are not mutated.

    Returns
    -------
    dict

    Raises
    ------
    ValueError
        If *object_type* is not a recognised key.

    Examples
    --------
    >>> cfg = get_default_config("Text", {"Text": "Hello", "Font Size": 24})
    >>> cfg["Object Type"]
    'Text'
    """
    if object_type not in default_configurations:
        raise ValueError(
            f"Unknown object type: {object_type!r}. "
            f"Supported types: {list(default_configurations.keys())}"
        )
    config = deepcopy(default_configurations[object_type])
    if overrides:
        config.update(overrides)
    return config


def apply_metadata(
    core_props: CoreProperties,
    metadata_dict: Dict[str, Any],
) -> None:
    """
    Apply document metadata to a Presentation's core properties.

    Recognised keys: ``"Author"``, ``"Category"``, ``"Comments"``,
    ``"Created"``, ``"Description"``, ``"Keywords"``,
    ``"Last Modified By"``, ``"Modified"``, ``"Subject"``, ``"Title"``.

    Unrecognised keys (e.g. ``"Filename"``, ``"Slide Aspect Ratio"``) are
    silently ignored — they are workflow config, not OOXML metadata.

    Parameters
    ----------
    core_props : CoreProperties
        ``prs.core_properties`` from a Presentation object.
    metadata_dict : dict
        Metadata key→value mapping.
    """
    _FIELDS = {
        "Author":          "author",
        "Category":        "category",
        "Comments":        "comments",
        "Created":         "created",
        "Description":     "description",
        "Keywords":        "keywords",
        "Last Modified By":"last_modified_by",
        "Modified":        "modified",
        "Subject":         "subject",
        "Title":           "title",
    }
    for key, value in metadata_dict.items():
        attr = _FIELDS.get(key)
        if attr and hasattr(core_props, attr):
            setattr(core_props, attr, value)


def show_functions() -> None:
    """Print all public functions and data objects in this module."""
    sections = {
        "Utilities": [
            "get_image_details(img_path)",
            "color_to_rgb(color)",
            "extract_ltwh(config)",
        ],
        "PPTX Object Functions": [
            "create_slide_deck(config_details, verbose=False)",
            "add_autoshape(slide, config)",
            "add_connector(slide, config)",
            "add_image(slide, config)",
            "add_notes(slide, note_text)",
            "add_shape_formatting(shape, config)",
            "add_table(slide, config)",
            "add_textbox(slide, config)",
            "add_text_formatting(run, config)",
        ],
        "Wrapper Functions": [
            "add_banners(slide, config)",
            "add_header(slide, config)",
        ],
        "Helper Functions": [
            "get_default_config(object_type, overrides=None)",
            "apply_metadata(core_props, metadata_dict)",
            "show_functions()",
            "show_autoshapes()",
            "show_object_alignment()",
            "show_dash_styles()",
            "show_slide_templates()",
        ],
        "Data": [
            "PPTX_LOOKUP          — enum lookup dict (align, valign, dash_styles, shapes, connectors)",
            "OBJECT_TYPE_HANDLERS — dispatch table for slide object rendering",
            "default_configurations — default config dicts for every object type",
        ],
    }
    print(f"\npptx_functions  v{__version__}")
    print("=" * 60)
    for section, items in sections.items():
        print(f"\n{section}:")
        for item in items:
            print(f"  - {item}")
    print('\nType help(<function_name>) for full parameter documentation.')


def show_autoshapes() -> None:
    """Print available auto shape key strings for use in ``"AutoShape Key"``."""
    print("Available AutoShape Keys:")
    for key in PPTX_LOOKUP["shapes"]:
        if key != "default":
            print(f"  - {key}")


def show_object_alignment() -> None:
    """Print available alignment key strings for use in ``"Align"``."""
    print("Available Alignment Keys:")
    for key in PPTX_LOOKUP["align"]:
        print(f"  - {key}")


def show_dash_styles() -> None:
    """Print available dash style key strings for use in ``"Line Style"``."""
    print("Available Dash Style Keys:")
    seen: set = set()
    for key, val in PPTX_LOOKUP["dash_styles"].items():
        if val not in seen:
            print(f"  - {key!r}")
            seen.add(val)


def show_slide_templates() -> None:
    """Print slide layout template descriptions."""
    print("Slide Templates:")
    for key, desc in _SLIDE_TEMPLATES.items():
        print(f"  {key}: {desc}")


################################################################################
# DEFAULT CONFIGURATIONS
################################################################################

default_configurations: Dict[str, Dict[str, Any]] = {

    "AutoShape": {
        "Add?":          True,
        "Object Type":   "AutoShape",
        "AutoShape Key": "rectangle",
        "left": 3.0,  "top": 3.0, "width": 4.0, "height": 2.5,
        "Fill Color":  "#6A0DAD", "Fill Alpha": 1.0,
        "Line Color":  "#FFA500", "Line Width": 2.5, "Line Style": "dash",
    },

    "Banner": {
        "Add?":        True,
        "Object Type": "Banner",
        "Text":        "",
        "left": 0.0,  "top": 0.0, "width": 10.0, "height": 0.4,
        "Align":       "center",
        "V Align":     "middle",
        "Font Name":   "Calibri",
        "Font Size":   14,
        "Font Color":  "#535353",
        "Bold?": False, "Italic?": False, "Underline?": False, "Word Wrap?": False,
        "Fill Color":  None,      "Fill Alpha": 1.0,
        "Line Color":  None,      "Line Width": 0, "Line Style": "-",
    },

    "Connector": {
        "Add?":        False,
        "Object Type": "Connector",
        "Type":        "straight",
        "Start X": 1.5, "Start Y": 1.0, "End X": 8.5, "End Y": 1.0,
        "Color":  "#535353", "Width": 2, "Style": "-",
    },

    "Image": {
        "Add?":        True,
        "Object Type": "Image",
        "img_path":    "",
        "Preserve Aspect Ratio?": True,
        "fit":         "width",
        "left": 3.25, "top": 1.5, "width": 5.0, "height": 5.0,
        "Line Width": 0, "Line Color": "#000000", "Line Style": "-",
    },

    "Table": {
        "Add?":           True,
        "Object Type":    "Table",
        "left": 1.0, "top": 1.5, "width": 8.0, "height": 4.0,
        "Columns":        3,
        "Rows":           1,
        "Column Widths":  [2.0, 3.0, 3.0],
        "Row Height":     0.4,
        "Column Headers": ["Column 1", "Column 2", "Column 3"],
        "Row Data":       [],
        "Font Size":      10,
        "Font Color":     "#333333",
        "Align":          "center",
        "V-Align":        "middle",
        "Bold?":          False,
    },

    "Text": {
        "Add?":        True,
        "Object Type": "Text",
        "Text":        "",
        "left": 1.5, "top": 1.0, "width": 10.0, "height": 4.0,
        "Align":       "center",
        "V Align":     "middle",
        "Font Name":   "Calibri",
        "Font Size":   18,
        "Font Color":  "#535353",
        "Bold?": False, "Italic?": False, "Underline?": False, "Word Wrap?": True,
        "Fill Color":  None, "Fill Alpha": 1.0,
        "Line Color":  None, "Line Width": 0, "Line Style": "-",
    },

    "Title": {
        "Add?":        True,
        "Object Type": "Text",
        "Text":        "Slide Title",
        "left": 1.5, "top": 0.3, "width": 8.5, "height": 0.5,
        "Align":       "left",
        "V Align":     "middle",
        "Font Name":   "Calibri",
        "Font Size":   28,
        "Font Color":  "#535353",
        "Bold?": False, "Italic?": False, "Underline?": False, "Word Wrap?": True,
        "Fill Color":  None, "Fill Alpha": 1.0,
        "Line Color":  None, "Line Width": 0, "Line Style": "-",
    },
}


################################################################################
# DISPATCH TABLE
################################################################################

OBJECT_TYPE_HANDLERS: Dict[str, Any] = {
    "AutoShape": add_autoshape,
    "Banner":    add_banners,
    "Connector": add_connector,
    "Header":    add_header,
    "Image":     add_image,
    "Table":     add_table,
    "Text":      add_textbox,
}
