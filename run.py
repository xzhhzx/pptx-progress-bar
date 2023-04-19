import collections.abc  # (workaround for python 3.10. See more: https://stackoverflow.com/questions/69468128/fail-attributeerror-module-collections-has-no-attribute-container)
import random
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.dml.color import RGBColor

PROGRESS_BAR_TAG = "progress_bar_tag"
CHAPTER_SLIDE_LAYOUT_NAME = "节标题"
CHAPTER_COLORS = ['540d6e', 'ee4266', 'ffd23f', '3bceac']

class ChapterColorsFactory(object):
    # Available colors for different chapters
    colors = CHAPTER_COLORS

    def __init__(self):
        self.ptr = 0

    def _convert_hex_to_rgb(self, hex_color):
        return (
            int(hex_color[:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:], 16)
        )

    def getCurrentColor(self):
        return RGBColor(*self._convert_hex_to_rgb(ChapterColorsFactory.colors[self.ptr]))

    def changeToNextColor(self):
        self.ptr = (self.ptr + 1) % len(ChapterColorsFactory.colors)

    def resetColor(self):
        self.ptr = 0


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
            if shape.name.startswith(PROGRESS_BAR_TAG):
                slide.shapes.element.remove(shape.element)


def addAllProgressBars(prs, chapter_tuple_list):
    chapterColorsFactory = ChapterColorsFactory()

    ptr = 0
    # Iterate through slides
    for current_page, slide in enumerate(prs.slides):

        progrss_bar_height = Inches(0.3)
        # Iterate the bar
        if chapter_tuple_list[ptr + 1][0] == current_page:
            # Move ptr
            ptr += 1

        print("====== current_page=", current_page, ", ptr=", ptr)
        end_page = 0

        # 1.Add multiple forgrounds
        for i in range(1, ptr + 1):
            start_page = chapter_tuple_list[i-1][0]
            end_page = chapter_tuple_list[i][0]
            num_pages = end_page - start_page
            delta_ratio = num_pages / len(prs.slides)
            offset_ratio = start_page / len(prs.slides)
            print("\t====== start_page=", start_page, ", end_page=", end_page)

            # Add rect
            slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                prs.slide_width * offset_ratio,
                prs.slide_height - progrss_bar_height,
                prs.slide_width * delta_ratio,
                progrss_bar_height
            )
            progress_bar = slide.shapes[-1]
            progress_bar.name = PROGRESS_BAR_TAG + '_fore_' + str(i) # tag a name for the shape
            progress_bar.fill.solid()
            progress_bar.line.fill.solid()
            progress_bar.fill.fore_color.rgb = chapterColorsFactory.getCurrentColor()
            progress_bar.line.color.rgb = chapterColorsFactory.getCurrentColor()
            chapterColorsFactory.changeToNextColor()


        # 2.Add foreground
        offset_ratio = end_page / len(prs.slides)
        delta_ratio = (current_page - end_page) / len(prs.slides)
        slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            prs.slide_width * offset_ratio,
            prs.slide_height - progrss_bar_height,
            prs.slide_width * delta_ratio,
            progrss_bar_height
        )
        print("Draw rect: ", prs.slide_width * offset_ratio, ", ", prs.slide_width * delta_ratio)
        progress_bar = slide.shapes[-1]
        progress_bar.name = PROGRESS_BAR_TAG + '_fore_current' # tag a name for the shape
        progress_bar.fill.solid()
        progress_bar.line.fill.solid()
        c1 = round(random.random() * 255)
        c2 = round(random.random() * 255)
        c3 = round(random.random() * 255)
        progress_bar.fill.fore_color.rgb = chapterColorsFactory.getCurrentColor()
        progress_bar.line.color.rgb = chapterColorsFactory.getCurrentColor()
        chapterColorsFactory.changeToNextColor()

        # # Add background
        # slide.shapes.add_shape(
        #     MSO_SHAPE.RECTANGLE,
        #     prs.slide_width * ratio,
        #     prs.slide_height - progrss_bar_height,
        #     prs.slide_width * (1 - ratio),
        #     progrss_bar_height
        # )
        # progress_bar_back = slide.shapes[-1]
        # progress_bar_back.name = PROGRESS_BAR_TAG + '_back' # tag a name for the shape

        chapterColorsFactory.resetColor()


# Instantiate a Presentation
path_to_presentation = "./test.pptx"
prs = Presentation(path_to_presentation)
# shapes = prs.slides[0].shapes
# printAttrs(shapes)
# print("\n_____________________________________________________\n")
# shapes = prs.slides[0].shapes[-1]
# printAttrs(shapes)


# Remove progress bars
removeAllProgressBars(prs.slides)

# Calculate the chapter segments
# Each tuple represents (<chapter_start_page_index>, <chapter_name>)
chapter_tuple_list = [(0, "start")]
for idx, slide in enumerate(prs.slides):
    if slide.slide_layout.name == CHAPTER_SLIDE_LAYOUT_NAME:
        chapter_tuple_list.append((idx, slide.shapes[0].text))
chapter_tuple_list.append((len(prs.slides), "end"))
print(chapter_tuple_list)

# Add progress bar
addAllProgressBars(prs, chapter_tuple_list)

# printSlidesStrcture(prs.slides)
prs.save(path_to_presentation)
