import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.dml.color import RGBColor

""" Default params """
# Tag name for progress bar, which is convenient for searching shapes on a slide by name
PROGRESS_BAR_TAG = "progress_bar_tag"
# The fixed name for the chapter slide of Microsoft Powerpoint (in Chinese).
CHAPTER_SLIDE_LAYOUT_NAME = "节标题"

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


class ProgressBar(object):
    """ A class that represents the progress bar template that can be drawn on
        a PPT slide. To use this, first instantiate with the builder. Then use
        the method *drawBarOnPage()* to draw the progress bar on a certain page.
    """
    def __init__(self, builder):
        self.position = builder.position
        self.thk = builder.thk
        # self.span = builder.span
        # self.colors = builder.colors
        self.chapterColorsFactory = builder.chapterColorsFactory
        self.chapterColorsFactoryBG = builder.chapterColorsFactoryBG
        # self.rect_layouts = builder.rect_layouts
        self.unit_size = builder.unit_size
        self.W = builder.W
        self.H = builder.H
        self.prs = builder.prs

        self.chapter_tuple_list = builder.chapter_tuple_list
        self.chapter_start_pages = [i[0] for i in self.chapter_tuple_list]
        self.num_pages_of_chapters = [i[0] - i[1] for i in zip(self.chapter_start_pages[1:], self.chapter_start_pages[:-1])]

    def _appendRect(self, shapes, offset, delta, chapterColorsFactory):
        """ Append a rectangle to the progress bar """
        if self.position in ['down', 'up']:
            rect = shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                offset,
                (self.H - self.thk) if self.position == 'down' else 0,
                delta,
                self.thk
            )
        elif self.position in ['right', 'left']:
            rect = shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                (self.W - self.thk) if self.position == 'right' else 0,
                offset,
                self.thk,
                delta
            )
        rect.fill.solid()
        rect.fill.fore_color.rgb = chapterColorsFactory.getCurrentColor()
        rect.line.fill.background()     # set line to transparent
        chapterColorsFactory.changeToNextColor()
        return rect

    def _drawBarOnPage(self, page):
        """ Draw a progress bar for a certain page on slide """
        slide = self.prs.slides[page]
        group_shape = slide.shapes.add_group_shape([])
        group_shape.name = PROGRESS_BAR_TAG  # tag a name for the shape
        ptr = 0
        offset = 0
        page += 1 # TODO: this is ugly. Try fixing the offset issue with an alternative method

        # 1.Previous full chapters
        while self.chapter_tuple_list[ptr+1][0] < page:
            delta = self.unit_size * self.num_pages_of_chapters[ptr]
            self._appendRect(group_shape.shapes, offset, delta, self.chapterColorsFactory)
            offset += delta
            ptr += 1

        # 2.Current chapter
        delta_pages = page - self.chapter_tuple_list[ptr][0]
        delta = self.unit_size * delta_pages
        self._appendRect(group_shape.shapes, offset, delta, self.chapterColorsFactory)
        offset += delta

        # 3.Remaining all pages
        total_pages = self.chapter_start_pages[-1]
        delta = self.unit_size * (total_pages - page)
        self._appendRect(group_shape.shapes, offset, delta, self.chapterColorsFactoryBG)
        offset += delta

        self.chapterColorsFactory.resetColor()
        self.chapterColorsFactoryBG.resetColor()

    def drawAllBars(self):
        """ Draw all progress bars for the whole presentation """
        for i in range(len(self.prs.slides)):
            self._drawBarOnPage(i)

    def removeAllBars(self):
        """ Remove all progress bars for the whole presentation """
        for slide in self.prs.slides:
            # Search shape named with "progress_bar" and remove
            for shape in slide.shapes:
                if shape.name.startswith(PROGRESS_BAR_TAG):
                    slide.shapes.element.remove(shape.element)

    class ProgressBarBuilder(object):
        """ Builder pattern"""
        def __init__(self, presentation):
            # Required params
            self.prs = presentation
            self.W = presentation.slide_width
            self.H = presentation.slide_height
            self.chapter_tuple_list = self._calculateChapterSegments()
            print(self.chapter_tuple_list)

            # Optional params
            self.position = 'down'
            self.thk = Inches(0.3)
            self.chapterColorsFactory = ChapterColorsFactory(["540d6e", "ee4266", "ffd23f", "3bceac"])
            self.chapterColorsFactoryBG = ChapterColorsFactory(["D8E1E9"])
            self.span = 'TODO'

        def _calculateChapterSegments(self):
            """ Pre-calculate all chapter segments.
                FORMAT: each tuple represents (<chapter_start_page_index>, <chapter_name>)
            """
            chapter_tuple_list = [(0, "start")] # (start page)
            for idx, slide in enumerate(self.prs.slides):
                if slide.slide_layout.name == CHAPTER_SLIDE_LAYOUT_NAME:
                    chapter_tuple_list.append((idx, slide.shapes[0].text))
            chapter_tuple_list.append((len(self.prs.slides), "end")) # (end_page)
            return chapter_tuple_list

        def setPosition(self, position):
            # TODOs: add validation check
            self.position = position
            return self

        def setThickness(self, thk):
            self.thk = Inches(thk)
            return self

        def setColors(self, colors):
            self.chapterColorsFactory = ChapterColorsFactory(colors)
            return self

        def setBgColors(self, bg_colors):
            self.chapterColorsFactoryBG = ChapterColorsFactory(bg_colors)
            return self

        # TODO
        def setSpan(self, span):
            self.span = span
            return self

        def build(self):
            self.x = (self.W - self.thk) if self.position == 'right' else 0
            self.y = (self.H - self.thk) if self.position == 'down' else 0
            self.num_pages = self.chapter_tuple_list[-1][0]
            if self.position in ['down', 'up']:
                self.unit_size = self.W / self.num_pages  # the base unit of the bar on one page
            elif self.position in ['right', 'left']:
                self.unit_size = self.H / self.num_pages
            return ProgressBar(self)
