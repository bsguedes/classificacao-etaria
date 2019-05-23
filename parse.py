import csv
from pptx.dml.color import RGBColor


class BonusCard:
    def __init__(self, name, text, range1, p1, range2, p2, score, perc, positions):
        self.name = name
        self.text = text
        self.range1 = range1 if len(range1) > 0 else None
        self.p1 = p1 if len(p1) > 0 else None
        self.range2 = range2 if len(range2) > 0 else None
        self.p2 = p2 if len(p2) > 0 else None
        self.score = score if len(score) > 0 else None
        self.perc = perc if len(perc) > 0 else None
        if len(positions) > 0:
            self.positions = [(int(x.split(';')[0]), float(x.split(';')[1])) for x in positions.split('&')]
        else:
            self.positions = None

    def type(self):
        return 'mult' if self.score is not None else 'range'

    def percent_text(self):
        return ("(%s das cartas)" % self.perc) if self.perc is not None else ''


class Card:
    def __init__(self, programming, tokens, name, attribute, year, points, age, money, ability, lore, img_path, gender,
                 positions):
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
        if len(positions) > 0:
            self.positions = [(int(x.split(';')[0]), float(x.split(';')[1])) for x in positions.split('&')]
        else:
            self.positions = None

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


def bonus_csv(file):
    cards = []
    with open(file, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        data = [r for r in reader]
        for c in data:
            cards.append(
                BonusCard(c['Nome'], c['Habilidade'], c['Range1'], c['ScoreRange1'], c['Range2'], c['ScoreRange2'],
                          c['ScoreRangeCard'], c['PercCE'], c['Tokens']))
    return cards


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
            positions = c['Tokens']
            cards.append(
                Card(programming, tokens, name, attribute, year, points, age, money, ability, lore, img_path, gender,
                     positions))
    return cards

