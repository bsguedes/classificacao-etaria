from pptx import Presentation
from pptx.util import Inches
from parse import parse_csv
from slides import fill_slide

presentation = Presentation()
presentation.slide_width = Inches(2.5)
presentation.slide_height = Inches(3.5)
cards = parse_csv('cards.csv')


def add_slide_from_card(card):
    layout = presentation.slide_layouts[6]
    slide = presentation.slides.add_slide(layout)
    fill_slide(slide, card)


for card in cards:
    add_slide_from_card(card)

presentation.save('cards.pptx')

