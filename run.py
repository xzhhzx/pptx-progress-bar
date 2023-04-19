import collections.abc  # (workaround for python 3.10. See more: https://stackoverflow.com/questions/69468128/fail-attributeerror-module-collections-has-no-attribute-container)
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.dml.color import RGBColor

PROGRESS_BAR_TAG = "progress_bar_tag"

# Debug util functions
def printAttrs(obj):
    accessibleAttrs = list(filter(lambda x: not x.startswith('_') ,dir(obj)))
    print("=== ", type(obj), " ===: ", accessibleAttrs, '\n')
    for attr in accessibleAttrs:
        try:
            print(attr, ": ", getattr(obj, attr))
        except Exception as e:
            print("=== error raised for attr: " + attr)

def printSlidesStrcture(slides):
    for idx, slide in enumerate(slides):
        print(idx, slide)
        for shape in slide.shapes:
            print("\t", shape)


# Helper functions
def removeAllProgressBars(slides):
    for idx, slide in enumerate(slides):
        # Search shape named with "progress_bar" and remove
        for shape in slide.shapes:
            if shape.name == PROGRESS_BAR_TAG:
                slide.shapes.element.remove(shape.element)

def addAllProgressBars(prs):
    for idx, slide in enumerate(prs.slides):
        progrss_bar_height = Inches(0.3)
        slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0,
            prs.slide_height - progrss_bar_height,
            prs.slide_width,
            progrss_bar_height
        )
        progress_bar = slide.shapes[-1]
        progress_bar.name = PROGRESS_BAR_TAG # tag a name for the shape
        progress_bar.fill.solid()
        progress_bar.fill.fore_color.rgb = RGBColor(100, 0, 0)
        progress_bar.line.color.rgb = RGBColor(100, 0, 0)


# Instantiate a Presentation
path_to_presentation = "./test.pptx"
prs = Presentation(path_to_presentation)
shapes = prs.slides[0].shapes
printAttrs(shapes)
print("\n_____________________________________________________\n")
shapes = prs.slides[0].shapes[-1]
printAttrs(shapes)

# Remove progress bars
removeAllProgressBars(prs.slides)

# Add progress bar
addAllProgressBars(prs)

# printSlidesStrcture(prs.slides)
prs.save(path_to_presentation)
