# Using python-pptx to Automate PowerPoint Slide Creation

**References:**
* [python-pptx Documentation](https://python-pptx.readthedocs.io/en/latest/index.html)

### This Python script uses the `python-pptx` library to automate the creation of PowerPoint presentations. This workflow uses abstraction and configuration dictionaries to create PowerPoint slide decks. The Python script currently supports adding:

* Auto Shapes
* Connectors
* Images
* Tables
* Text Boxes

### Other preformatted objects can be added as well within this workflow:

* Slide Titles
* Banners
* Bulleted Lists

Users can define the Presentation Aspect Ratio  (On-Screen (4:3) or Widescreen (16:9)) and add metadata to the slide via the python script. Slide layouts, object properties, and metadata can be defined in the `slide_config` dictionary. The python script processes these configurations to place objects on the slide consistent the user-defined formatting, saving the result as a `.pptx` file.

* Author
* Created
* Description
* Title
* Subject
* Comments
* Keywords
* Category
*  Last Modified by
*  Modified

The `slide_config` dictionary can be modified to add slides and object configurations leveraging the `get_default_config` function for a quick setup. For example:

```python
    
    {"Slide Template": 6, ## Add a blank slide to the Presentation 

     ## Add a Title - a formatted Text object
     "Title Config": get_default_config("Title", {"text": "Professor Hubert J. Farnsworth"}),

     ## Add an Image
     "Image Config": get_default_config("Image", {"img_path": "img/FuturamaProfessorFarnsworth.png"})}
    
```
which creates the slide below:

<center><img src="img/FuturamaProfessorFarnsworth.png"></center>

The configuration dictionary follows the format of:

```python

slide_config = {"Details": {"Author": "Bender",
                            "Created": "08 April 2025",
                            "Description": "Example PowerPoint slide configuration.",
                            "Title": "Workflow Automation Example",
                            "Subject": "Weekly Update",
                            "Comments": "Generated programmatically using python-pptx",
                            "Keywords": "python_pptx, lorem, webcolors, PIL, dateutil, re, pandas, copy",
                            "Category": "Workflow Automation",
                            
                            "Filename": "Filename_of_Slide-Deck.pptx",     ## Filename of Created PowerPoint Slide Deck
                            "Slide Aspect Ratio": "4:3"},                  ## On-Screen Layout / Widescreen is "16:9"

                "Slides": {"Slide 01": {"Slide Template": 6, ## Add a blank slide to the Presentation
                                        "Slide Name": "Cover Slide",

                                        "Objects": {"Text Config": "...",
                                                    "Image Config": "..."},

                           "Slide 02": {"Slide Template": 6, ## Add a blank slide to the Presentation
                                        "Slide Name": "Introduction",


                                        "Objects": {"Text Config": "...",
                                                    "Image Config": "..."},

                                                    ## Add two Image objects
                                                    "Image 01 Config": "...",
                                                    "Image 02 Config": "..."}
                                       }
               }
            
```

creates a slide with a title in the title placeholder and a bulleted list in the content placeholder. The main loop iterates over this dictionary, calling specific functions (e.g., `add_title_to_placeholder`, `add_bulleted_list_to_placeholder`) to add each object, applying defaults or custom overrides as specified.

The script is modular and extensible—users can run it as-is for a basic presentation or customize configs for specific needs, like adding a red connector with 

```python
    get_default_config("Connector", {"Color": "#FF0000"})
```    

It’s ideal for repetitive slide generation and requires only Python and several other libraries --`python-pptx`, `webcolors`, `PIL`, installed, and a basic understanding of the config structure to get started.

There are several helper functions that display the allowed configuration values for formatting various pptx objects:

```python
    ## Lists the allowed Auto Shape keys in this workflow for the PPTX Auto Shape Objects
    show_autoshapes():
    
    ## Lists the allowed Auto Shape keys in this workflow for the PPTX Auto Shape Objects
    show_object_alignment():

    ## Lists the allowed Line Dash Styles allowed when formatting PPTX Line Objects
    show_dash_styles():
```
