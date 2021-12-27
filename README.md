# Excel Report Evaluator

Simple Python GUI with Tkinter for evaluating Microsoft Excel reports (.xlsx-Files).

At this point the approach is based on the cell background color. The GUI allows to open a local directory and will add all Excel files to a list.
When processing the reports the chosen color will determine wether a report will pass (=not found) or not pass (=found).

![Main GUI with Color Chooser](https://raw.githubusercontent.com/yet-another-alex/excel-report-evaluator/main/screens/screen1.png)

Many automated reports have multiple hundreds of lines and are already configured to set the background color based on its status. This simple script can aid with the bulk evaluation.

![Main GUI after report processing](https://raw.githubusercontent.com/yet-another-alex/excel-report-evaluator/main/screens/screen2.png)

In the above example test2 and test3 have no issues (meaning no red cells) and test1 has at least some issues (meaning some red cells).

Planned features are currently to extract the highlighted information and evaluate on a rule-based approach instead of by cell background color only.
