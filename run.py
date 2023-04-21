import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.dml.color import RGBColor

from progress_bar import ProgressBar

# Fixed params
PROGRESS_BAR_TAG = "progress_bar_tag"
CHAPTER_SLIDE_LAYOUT_NAME = "节标题"
IS_DEBUG = True

# Adjustable params
config = yaml.safe_load(open('./config.yaml'))
print(config)
progress_bar_thickness = Inches(config['progress_bar']['thickness'])
# chapterColorsFactory = ChapterColorsFactory(config['progress_bar']['colors'])


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


def addAllProgressBars(prs, chapter_tuple_list):
    pb = ProgressBar.ProgressBarBuilder(prs.slide_width, prs.slide_height, chapter_tuple_list) \
        .setPosition('left') \
        .setColors(['123456', 'b8c9a0']) \
        .setThickness(0.4) \
        .build()

    for i in range(0, chapter_tuple_list[-1][0]):
        pb.drawBarOnPage(i, prs.slides[i])
        # for sh in prs.slides[i].shapes:
        #     print(i, ": ", sh.name)
    return


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
