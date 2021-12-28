from openpyxl import load_workbook
import os

COLOR_ERROR = '#8b0000'
COLOR_OK = '#006400'

class Report:
    def __init__(self, name, result, details):
        self.name = name
        self.result = result
        self.details = details


def evaluate(v1, v2, operator):
    if operator == "==":
        return v1 == v2
    elif operator == "<=":
        return v1 <= v2
    elif operator == ">=":
        return v1 >= v2
    elif operator == "!=":
        return v1 != v2
    elif operator == ">":
        return v1 > v2
    elif operator == "<":
        return v1 < v2

def report_thread(queue, filename, working_directory, rule):
    wb = load_workbook(os.path.join(working_directory, filename))

    lines = list()
    result = Report(filename, "", "")

    # iterate the rows
    found = False
    for line in wb.active.iter_rows():
        for c in line:
            if rule.type == "BGCOLOR":
                # check if the color is indexed for legacy colors
                col = ''
                if c.fill.bgColor.type == 'indexed':
                    col = c.fill.fgColor.rgb
                else:
                    col = c.fill.bgColor.rgb
                
                if col == rule.compare:
                    found = True
            elif rule.type == "FONTCOLOR":
                if c.font.color.rgb == rule.compare:
                    found = True
            elif rule.type == "VALUE" and (type(c.value) == int or type(c.value) == float):
                found = evaluate(c.value, rule.compare, rule.operator)
            elif rule.type == "COLVALUE" and (type(c.value) == int or type(c.value) == float):
                if c.column == rule.col:
                    found = evaluate(c.value, rule.compare, rule.operator)
        
        # search criteria was found
        if found:
            new_line = ""
            for c in line:
                new_line += str(c.value)
                if c != line[-1]:
                    new_line += ", "
            new_line += "\n"
            lines.append(new_line)
            found = False
        
    if len(lines) == 0:
        result.result = "OK"
    else:
        result.result = "ERROR"
        # save error lines
        for line in lines:
            result.details += line
            result.details += "\n"

    lines.clear()
    queue.put(result)