import csv
from pptx.dml.color import RGBColor


class Card:
    def __init__(self, programming, tokens, name, attribute, year, points, age, money, ability, lore, img_path, gender):
        self.programming = programming
        self.tokens = tokens
        self.name = name
        self.attribute = attribute
        self.year = year
        self.points = points
        self.age = age
        self.money = money
        self.ability = ability
        self.lore = lore
        self.img_path = img_path
        self.gender = gender

    def ability_text(self):
        if self.ability[0] == 'Nada':
            return ''
        else:
            return "%s: %s" % (self.ability[0].upper(), self.ability[1])

    def ability_color(self):
        if self.ability[0] == 'Quando Ativado':
            return RGBColor(241, 194, 50)
        elif self.ability[0] == 'Permanente':
            return RGBColor(217, 210, 233)
        elif self.ability[0] == 'Uma Vez Entre Turnos':
            return RGBColor(244, 204, 204)
        elif self.ability[0] == 'Polêmico':
            return RGBColor(224, 102, 102)
        else:
            return RGBColor(255, 255, 255)


def pint(value):
    return float(value) if value != '' else 0


def parse_csv(file):
    cards = []
    with open(file, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        data = [r for r in reader]
        for c in data[2:]:
            programming = (c['A1'] == 'TRUE', c['A2'] == 'TRUE', c['A3'] == 'TRUE')
            tokens = [pint(c['X1']), pint(c['X2']), pint(c['X3']), pint(c['X4']), pint(c['X5']), pint(c['X'])]
            name = c['Name']
            attribute = c['Label']
            year = c['Year']
            points = c['Points']
            age = c['Age']
            money = c['Cash']
            ability = (c['AbType'], c['AbText'])
            img_path = c['ImgPath']
            gender = c['Gender']
            lore = c['Lore']
            cards.append(
                Card(programming, tokens, name, attribute, year, points, age, money, ability, lore, img_path, gender))
    return cards

