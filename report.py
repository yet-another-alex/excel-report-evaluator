from openpyxl import load_workbook
import os

COLOR_ERROR = '#8b0000'
COLOR_OK = '#006400'


class Report:
    def __init__(self, name, result, details):
        self.name = name
        self.result = result
        self.details = details


def report_thread(queue, filename, working_directory, search_color):
    wb = load_workbook(os.path.join(working_directory, filename))

    lines = list()
    result = Report(filename, "", "")

    # iterate the rows
    found = False
    for line in wb.active.iter_rows():
        for c in line:
            if c.fill.bgColor.rgb == search_color:
                # if the search criteria for an error is found
                found = True
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