import csv


class Card:
    def __init__(self, programming, tokens, name, attribute, year, points, age, money, ability, lore):
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


def parse_csv(file):
    with open(file, newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=' ', quotechar='|')

    for row in spamreader:
        print(', '.join(row))
