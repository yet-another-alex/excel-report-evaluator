import threading
import queue
import logging
from tkinter import *
from tkinter import filedialog, ttk, messagebox
from rules import Rules
import report as rp
import os


def openfolder():
    """Callback-function for the "choose directory"-Button.
    Opens a filedialog and adds all Excel files to the excelfiles-list.
    Also updates the listbox.
    """
    filename = filedialog.askdirectory()
    if filename:
        global working_directory
        working_directory = filename
        for file in os.listdir(filename):
            if file.endswith('.xlsx'):
                excelfiles.append(file)
        for xl in excelfiles:
            listbox.insert('end', xl)
        statuslabel.config(text=str(len(excelfiles)) + ' reports loaded!')
        button_scan.config(state='normal')


def selection_changed(event):
    """Callback-function for the selection changed event of the listbox.
    When a user selects a different Report in the list, this function is called and will 
    update the outputText widget accordingly.

    Args:
        event (obj): Tkinter widget that sends the event (listbox)
    """
    # check selected item
    selection = listbox.curselection()
    if selection:
        selected_report = str(listbox.selection_get())
        statuslabel.config(text=selected_report)

        if(len(reports) > 0):
            matching_report = next(r for r in reports if r.name == selected_report)
            outputText.delete(1.0, END)

            if matching_report and len(matching_report.details) > 0:
                outputText.insert(1.0, matching_report.details)


def combo_selection_changed(event):
    """Callback-function for the selection changed event of the combobox.
    When a user changes the selected rule in the combobox, this function updates the label 
    that displays the details.

    Args:
        event (obj): Tkinter widget that sends the event (combobox)
    """
    selected_rule = next(r for r in eval_rules.ruleset if r.name == combobox_rules.get())
    rulelabel.configure(text=f"{selected_rule.type} {selected_rule.operator} {selected_rule.compare} (col: {selected_rule.col})")


def scan_thread():
    """Helper-function to update the GUI and keep the user from starting multiple scans.
    Also starts a second thread to actually process the reports.
    """
    statuslabel.config(text='evaluating .. reports will change color after processing!')
    button_scan.config(state='disabled')
    button_open.config(state='disabled')
    button_reset.config(state='disabled')
    
    t = threading.Thread(target=scan_reports)
    t.start()


def scan_reports():
    """Report evaluation happens here.
    The selected rule is retrieved and all Excel files will be evaluated accordingly in their own thread.
    All reports will be added to reports.
    The results will be displayed in the GUI via evaluate_results().
    """
    queue_reports = queue.Queue()
    thread_list = list()
    results = list()

    # get evaluation type
    selected_rule = next(r for r in eval_rules.ruleset if r.name == combobox_rules.get())

    for excel_file in excelfiles:
        report_eval = threading.Thread(target=rp.report_thread, args=(queue_reports, excel_file, working_directory, selected_rule))
        report_eval.start()
        thread_list.append(report_eval)

    for t in thread_list:
        t.join()

    while not queue_reports.empty():
        result = queue_reports.get()
        results.append(result)
        reports.append(result)
    
    evaluate_results(results=results)


def evaluate_results(results):
    """Displays the results of a report evaluation to the user via GUI.
    This function changes the background color of the listbox item containing the 
    report according to their result parameter.
    Also enables the reset-Button.

    Args:
        results (list of Report): List of the reports containing the results
    """
    for index, item in enumerate(listbox.get(0, END)):
        found_report = next(r for r in results if r.name == item)
        if found_report.result == 'OK':
            listbox.itemconfig(index, bg='green', selectbackground=rp.COLOR_OK)
        elif found_report.result == 'ERROR':
            listbox.itemconfig(index, bg='red', selectbackground=rp.COLOR_ERROR)

    button_reset.configure(state='normal')


def reset():
    """Simple helper function to reset the GUI to it's original state.
    """
    # reset GUI
    listbox.delete(0, END)
    button_open.configure(state='normal')
    button_scan.configure(state='disabled')
    statuslabel.configure(text=' .. reset!')
    outputText.delete(1.0, END)
    # reset data
    excelfiles.clear()
    reports.clear()

# MAIN script
# create root Tkinter element
root = Tk()

# logging configuration
logging.basicConfig(filename='errors.log', encoding='utf-8', level=logging.ERROR)

excelfiles = list()
reports = list()
global working_directory

# combobox for rules
combobox_rules = ttk.Combobox(root, state='readonly')
combobox_rules.grid(row=0, column=0, columnspan=2, padx='5', pady='5', sticky=W)
# load rules.json
try:
    eval_rules = Rules('rules.json')
    eval_rules.load()
    combobox_rules.configure(values=eval_rules.ruleset)
except Exception as err:
    messagebox.showerror('error while loading', f"Error while loading rules.json {err}")
    logging.exception(err)
    root.quit()

combobox_rules.bind('<<ComboboxSelected>>', combo_selection_changed)
combobox_rules.current(0)

# define status label for info to user
statuslabel = Label(root, text='waiting for reports...')
statuslabel.grid(row=8, column=0, columnspan=5, padx='5', pady='5', sticky=W)

# rule label
rulelabel = Label(root)
rulelabel.grid(row=0, column=2, columnspan=5, padx='5', pady='5', sticky=W)

# trigger combobox update
combo_selection_changed(None)

# define listbox for reports
listbox = Listbox(root, height=15, width=50, selectmode='single')
listbox.grid(row=1, column=0, rowspan=4, columnspan=5, padx='5', pady='5', sticky=NSEW)
listbox.bind('<<ListboxSelect>>', selection_changed)

# buttons for user interaction - BASE actions
button_open = Button(root, text='Choose directory', command=openfolder)
button_open.grid(row=7, column=0, columnspan=2, padx='5', pady='5')
button_quit = Button(root, text='Quit', command=root.quit)
button_quit.grid(row=8, column=4, padx='5', pady='5')
button_reset = Button(root, text='Reset', command=reset)
button_reset.grid(row=7, column=4, padx='5', pady='5')
button_scan = Button(root, text='Start evaluation', command=scan_thread, state='disabled')
button_scan.grid(row=7, column=2, columnspan=2, padx='5', pady='5')


outputlabel = Label(root, text='Errors in the following lines:')
outputlabel.grid(row=0, column=5, padx='5', pady='5', sticky=NSEW)
outputText = Text(root, height=20, width=60)
outputText.grid(row=1, column=5, rowspan=8, padx='5', pady='5', sticky=NSEW)

if __name__ == '__main__':
    # app settings
    root.title('Excel Report Evaluator')
    # prevent resizing
    root.resizable(False, False)

    root.mainloop()
