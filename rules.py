import json


class Rule:
    def __init__(self, name, type, compare, operator, col = -1):
        self.name = name
        self.operator = operator
        self.compare = compare
        self.type = type
        self.col = col

    def __str__(self):
        return self.name


class Rules:
    def __init__(self, filename):
        self.filename = filename
        self.ruleset = list()

    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__, 
        sort_keys=True, indent=4)

    def save(self):
        f = open(self.filename, "w")
        f.write(self.toJSON())
        f.close()

    def load(self):
        self.ruleset.clear()

        f = open(self.filename, "r")
        data = json.load(f)
        for el in data["ruleset"]:
            element = Rule(el["name"], el["type"], el["compare"], el["operator"], el["col"])
            self.ruleset.append(element)
        f.close()
