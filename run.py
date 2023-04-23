import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.dml.color import RGBColor

from progress_bar import ProgressBar
# from utils import *

IS_DEBUG = True

""" Adjustable params """
config = yaml.safe_load(open('./config.yaml'))
print(config)
progress_bar_thickness = config['progress_bar']['thickness']


if __name__ == '__main__':
    # Instantiate a Presentation
    path_to_presentation = "./test.pptx"
    prs = Presentation(path_to_presentation)
    pb = ProgressBar.ProgressBarBuilder(prs) \
        .setPosition('down') \
        .setColors(['c93456', '18c9a0', 'a2418a']) \
        .setThickness(0.2) \
        .build()
        # .setShapeName(PROGRESS_BAR_TAG) \ # TODO: add param

    # Clear all progress bars
    pb.removeAllBars()

    # Add progress bar
    pb.drawAllBars()

    prs.save(path_to_presentation)
    print(f"Run successful! File saved to {path_to_presentation} !")
