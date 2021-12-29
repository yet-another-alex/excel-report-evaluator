import json

class Rule:
    """Rule class containing all parameters required to define a rule.
    """

    def __init__(self, name, type, compare, operator, col = -1):
        """Initialization-function for the Rule-class.

        Args:
            name (string): name of the rule
            type (string): type of the rule, see readme.md for valid types
            compare (alphanumeric): value to compare the cell value to
            operator (string): operator used for comparison, see readme.md for valid operators
            col (int, optional): Column restraint. Starts at 1 for column A in Excel. Defaults to -1.
        """
        self.name = name
        self.operator = operator
        self.compare = compare
        self.type = type
        self.col = col

    def __str__(self):
        """Overwrites __str__ function to represent a Rule object by its name.

        Returns:
            string: self.name. Name of the rule.
        """
        return self.name


class Rules:
    """Rules class containing a filename and a ruleset (=list of Rule objects).
    Also contains functions for loading from a json-file as well as saving, although unused.
    """

    def __init__(self, filename):
        """Initialization function for the Rules class. Accepts a filename and initializes the Rule-list.

        Args:
            filename (string): name of the file to load
        """
        self.filename = filename
        self.ruleset = list()

    def toJSON(self):
        """Quick and dirty JSON conversion of this class using __dict__.

        Returns:
            string: Rules class representation as JSON
        """
        return json.dumps(self, default=lambda o: o.__dict__, 
        sort_keys=True, indent=4)

    def save(self):
        """Writes the Rules class to file (self.filename) as JSON.
        """
        f = open(self.filename, 'w')
        f.write(self.toJSON())
        f.close()

    def load(self):
        """Loads a JSON file and fills the Rules class including ruleset.
        """
        self.ruleset.clear()

        f = open(self.filename, 'r')
        data = json.load(f)
        for el in data['ruleset']:
            element = Rule(el['name'], el['type'], el['compare'], el['operator'], el['col'])
            self.ruleset.append(element)
        f.close()
