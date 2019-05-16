from pptx import Presentation
from pptx.util import Inches, Pt
from parse import parse_csv

presentation = Presentation()
presentation.slide_width = Inches(2.5)
presentation.slide_height = Inches(3.5)
cards = parse_csv('cards.csv')


def add_slide_from_card(card):
    layout = presentation.slide_layouts[6]
    presentation.slides.add_slide(layout)
    return None


for card in cards:
    add_slide_from_card(card)

presentation.save('cards.pptx')

