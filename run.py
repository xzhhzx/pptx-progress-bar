import yaml
from pptx import Presentation

from progress_bar import ProgressBarTemplate


# TODO: retrieve params from config.yaml
""" Adjustable params """
config = yaml.safe_load(open('./config.yaml'))
print('================== CONFIGURATION ==================')
print(yaml.dump(config))


if __name__ == '__main__':
    # Instantiate a Presentation
    path_to_presentation = config['presentation_path']
    prs = Presentation(path_to_presentation)

    # Create a ProgressBarTemplate
    pbt = ProgressBarTemplate.ProgressBarTemplateBuilder(prs) \
        .setPosition(config['progress_bar']['position']) \
        .setThickness(config['progress_bar']['finished_part']['thickness']) \
        .setColors(config['progress_bar']['finished_part']['colors']) \
        .setBgColor(config['progress_bar']['unfinished_part']['color']) \
        .setBgThicknessRatio(config['progress_bar']['unfinished_part']['relative_thickness_ratio']) \
        .setAddCaption(config['progress_bar']['finished_part']['add_caption']) \
        .build()
        # TODO: add more configurable params
        # .setPageMarginXY(0.15, 0.25) \
        # .setChapterBarShape('rounded_rectangle')
        # .setIfAddText(True)
        # .setIfDetectChapter(False)


    # Clear all progress bars
    pbt.removeAllBars()

    # Add progress bar
    pbt.drawAllBars()

    # TODO: add option not to overwrite
    prs.save(path_to_presentation)
    print(f"Run successful! File saved to {path_to_presentation} !")
