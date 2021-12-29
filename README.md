# Excel Report Evaluator

Simple Python GUI with Tkinter for evaluating Microsoft Excel reports (.xlsx-Files).

## Usage

Start *main.py* and choose one of the example rules that are currently included via the dropdown at the top left corner.

![Application started](https://raw.githubusercontent.com/yet-another-alex/excel-report-evaluator/main/screens/screen1.png)

The rules will be explained in more detail below. The default is looking for a red background in any Excel-cell.

Now click "Choose directory" and select a directory that contains .xlsx-files for evaluation.

![Reports were added](https://raw.githubusercontent.com/yet-another-alex/excel-report-evaluator/main/screens/screen2.png)

All .xlsx-files will be added to the list. Click "Start evaluation" to process the reports. Depending on file size this may take a while.

![Reports are processed](https://raw.githubusercontent.com/yet-another-alex/excel-report-evaluator/main/screens/screen3.png)

When the reports are processed the background color of the report in the list will change according to the selected ruleset.
In the above example, test2 did not contain any cells with a red background, but test1 and test3 did.
To see more details and extract the rows that contain the searched parameters, click on the report in the list.

![Found lines are extracted and displayed](https://raw.githubusercontent.com/yet-another-alex/excel-report-evaluator/main/screens/screen4.png)

When a report is flagged according to the specified rules, the affected lines will be displayed on the right in the Text widget.

![more example extracted lines](https://raw.githubusercontent.com/yet-another-alex/excel-report-evaluator/main/screens/screen5.png)

## JSON rule description

The sample json file provided **rules.json** is an example for every possible rule currently implemented.
The json file always needs to be named "rules.json" for now. Within the ruleset-list you can add, adjust or delete any of the components to your liking.

Every rule is required to have several parameters.

| parameter      | purpose       |
| ------------- | ------------- |
| col           | Column to use for evaluation. -1 means all columns will be evaluated.  |
| compare  | The value that is being compared to the cell value within the Excel file. For Color this is an Excel-compatible Color Code, for values this is an alphanumeric value.  |
| name  | The name of the rule that will be displayed in the GUI.  |
| operator | The operator that will be used when comparing the cell value with the *compare*-value. Currently supports **==, >=, <=, !=, <, >** .  |
| type  | The type of rule. See below for more information.  |


An example JSON-description of the rule looking for a red background within a cell would look like this:

        {
            "col": -1,
            "compare": "FFFF0000",
            "name": "RED BG",
            "operator": "==",
            "type": "BGCOLOR"
        }

If you wanted to modify this example to only look for a red background in column 7, it could look like this:

        {
            "col": 7,
            "compare": "FFFF0000",
            "name": "RED BG C7",
            "operator": "==",
            "type": "BGCOLOR"
        }
        

### Rule Types

| Rule Type  | description |
| ------------- | ------------- |
| BGCOLOR  | Looks for the background color of a cell.  |
| FONTCOLOR  | Looks for the color of the font within a cell.  |
| VALUE  | Looks to compare the cell value to the compare value based on the operator provided.  |
| COLVALUE | Same as value, but only in the specified column.  |
