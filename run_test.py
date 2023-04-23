import yaml
from pptx import Presentation

from progress_bar import ProgressBarTemplate


if __name__ == '__main__':
    # Instantiate a Presentation
    path_to_presentation = "./test.pptx"
    prs = Presentation(path_to_presentation)

    # Create a ProgressBarTemplate
    pbt = ProgressBarTemplate.ProgressBarTemplateBuilder(prs) \
        .setPosition('bottom') \
        .setThickness(0.2) \
        .setColors(['c93456', '18c9a0', 'a2418a']) \
        .setBgColor('D8E1E9') \
        .setBgThicknessRatio(0.75) \
        .build()
        # TODO: add more configurable params
        # .setPageMarginXY(0.15, 0.25) \
        # .setChapterBarShape('rounded_rectangle')
        # .setIfAddText(True)

    # Clear all progress bars
    pbt.removeAllBars()

    # Add progress bar
    pbt.drawAllBars()

    prs.save(path_to_presentation)
    print(f"Run successful! File saved to {path_to_presentation} !")
