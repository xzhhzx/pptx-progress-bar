import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.dml.color import RGBColor

# Fixed params
PROGRESS_BAR_TAG = "progress_bar_tag"
CHAPTER_SLIDE_LAYOUT_NAME = "节标题"
IS_DEBUG = True

# Adjustable params
config = yaml.safe_load(open('./config.yaml'))
print(config)
progress_bar_thickness = Inches(config['progress_bar']['thickness'])
# exit(0)

class ChapterColorsFactory(object):
    """ This class is for retrieving chapter colors. """
    def __init__(self, colors):
        self.ptr = 0
        self.colors = colors    # available colors (in hex format) for different chapters

    def _convert_hex_to_rgb(self, hex_color):
        return (
            int(hex_color[:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:], 16)
        )

    def getCurrentColor(self):
        return RGBColor(*self._convert_hex_to_rgb(self.colors[self.ptr]))

    def changeToNextColor(self):
        self.ptr = (self.ptr + 1) % len(self.colors)

    def resetColor(self):
        self.ptr = 0


# Util functions for debugging
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
def appendColoredRectangleToSlide(slideShapes, x, y, w, h, name, chapterColorsFactory):
    progress_bar = slideShapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    progress_bar.name = name # tag a name for the shape
    progress_bar.fill.solid()
    progress_bar.fill.fore_color.rgb = chapterColorsFactory.getCurrentColor()
    progress_bar.line.fill.background()     # set line to transparent
    chapterColorsFactory.changeToNextColor()


def addAllProgressBars(prs, chapter_tuple_list):
    chapterColorsFactory = ChapterColorsFactory(config['progress_bar']['colors'])
    ptr = 0

    # Iterate through slides
    for current_page, slide in enumerate(prs.slides):
        # If current page reaches next item in chapter_tuple_list (which means
        # reaches a new chapter), then move ptr
        if current_page == chapter_tuple_list[ptr + 1][0]:
            ptr += 1

        if IS_DEBUG: print("====== current_page=", current_page, ", ptr=", ptr)

        # 1. Add progress bars of previous chapters
        for i in range(1, ptr + 1):
            start_page = chapter_tuple_list[i-1][0]
            end_page = chapter_tuple_list[i][0]
            num_pages = end_page - start_page
            delta_ratio = num_pages / len(prs.slides)
            offset_ratio = start_page / len(prs.slides)
            if IS_DEBUG: print("\t====== start_page=", start_page, ", end_page=", end_page)

            # Append rectangle to the tail of the shape tree
            appendColoredRectangleToSlide(
                slide.shapes,
                prs.slide_width * offset_ratio,
                prs.slide_height - progress_bar_thickness,
                prs.slide_width * delta_ratio,
                progress_bar_thickness,
                PROGRESS_BAR_TAG + '_fore_' + str(i), # tag a name for the shape
                chapterColorsFactory
            )

        # 2. Add progress bar of current chapter (presented part)
        end_page = chapter_tuple_list[ptr][0]
        offset_ratio = end_page / len(prs.slides)
        delta_ratio = (current_page - end_page + 1) / len(prs.slides)
        appendColoredRectangleToSlide(
            slide.shapes,
            prs.slide_width * offset_ratio,
            prs.slide_height - progress_bar_thickness,
            prs.slide_width * delta_ratio,
            progress_bar_thickness,
            PROGRESS_BAR_TAG + '_fore_current', # tag a name for the shape
            chapterColorsFactory
        )
        if IS_DEBUG: print("Draw rect: ", prs.slide_width * offset_ratio, ", ", prs.slide_width * delta_ratio)

        # 3. Add progress bars background
        offset_ratio = (current_page + 1) / len(prs.slides)
        delta_ratio = 1 - offset_ratio
        background_bar_thickness = progress_bar_thickness * config['background']['relative_thickness_ratio']
        appendColoredRectangleToSlide(
            slide.shapes,
            prs.slide_width * offset_ratio,
            prs.slide_height - progress_bar_thickness / 2 - background_bar_thickness / 2,
            prs.slide_width * delta_ratio,
            background_bar_thickness,
            PROGRESS_BAR_TAG + '_back', # tag a name for the shape
            ChapterColorsFactory([config['background']['color']])
        )

        # Reset color after each slide
        chapterColorsFactory.resetColor()

def removeAllProgressBars(slides):
    for idx, slide in enumerate(slides):
        # Search shape named with "progress_bar" and remove
        for shape in slide.shapes:
            if shape.name.startswith(PROGRESS_BAR_TAG):
                slide.shapes.element.remove(shape.element)

if __name__ == '__main__':
    # Instantiate a Presentation
    path_to_presentation = "./test.pptx"
    prs = Presentation(path_to_presentation)

    # Clear all progress bars
    removeAllProgressBars(prs.slides)

    # Pre-calculate all chapter segments
    # Each tuple represents (<chapter_start_page_index>, <chapter_name>)
    chapter_tuple_list = [(0, "start")] # (start page)
    for idx, slide in enumerate(prs.slides):
        if slide.slide_layout.name == CHAPTER_SLIDE_LAYOUT_NAME:
            chapter_tuple_list.append((idx, slide.shapes[0].text))
    chapter_tuple_list.append((len(prs.slides), "end")) # (end_page)
    if IS_DEBUG: print(chapter_tuple_list)

    # Add progress bar
    addAllProgressBars(prs, chapter_tuple_list)

    prs.save(path_to_presentation)
    # printSlidesStrcture(prs.slides)
    print(f"Run successful! File saved to {path_to_presentation} !")
