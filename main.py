import threading
import queue
from tkinter import *
from tkinter import filedialog, colorchooser
import report as rp
import os


def openfolder():
    filename = filedialog.askdirectory()
    if filename:
        global working_directory
        working_directory = filename
        for file in os.listdir(filename):
            if file.endswith(".xlsx"):
                excelfiles.append(file)
        for xl in excelfiles:
            listbox.insert("end", xl)
        statuslabel.config(text=str(len(excelfiles)) + " reports loaded!")
        button_scan.config(state='normal')


def choose_color():
    global color_search
    color = colorchooser.askcolor(parent=root, title="Select search criteria color")
    colorLabel.configure(bg=color[1])
    color_search = color[1].replace("#", "FF")                                          # currently required for actual HEX values to be correctly transferred to MS Excel aRGB
    color_result_rgb = ' '.join(str(int(x)) for x in color[0])
    print(color_result_rgb) 
    print(color[0])
    print(color_search)

def selection_changed(event):
    # check selected item
    selection = listbox.curselection()
    if selection:
        selected_report = str(listbox.selection_get())
        statuslabel.config(text=selected_report)

        if(len(reports) > 0):
            matching_report = next(r for r in reports if r.name == selected_report)

            if matching_report:
                print(matching_report.details)


def scan_thread():
    statuslabel.config(text="evaluating .. reports will change color after processing!")
    button_scan.config(state='disabled')
    button_open.config(state='disabled')
    button_reset.config(state='disabled')

    t = threading.Thread(target=scan_reports)
    t.start()


def scan_reports():
    queue_reports = queue.Queue()
    thread_list = list()
    results = list()

    for excel_file in excelfiles:
        report_eval = threading.Thread(target=rp.report_thread, args=(queue_reports, excel_file, working_directory, color_search))
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
    i = 0
    while i < listbox.size():
      for report in results:
        if report.result == "OK":
            listbox.itemconfig(i, bg='green', selectbackground=rp.COLOR_OK)
        elif report.result == "ERROR":
            listbox.itemconfig(i, bg='red', selectbackground=rp.COLOR_ERROR)
        i += 1
    
    button_reset.configure(state='normal')


def reset():
    # reset GUI
    listbox.delete(0, END)
    button_open.configure(state='normal')
    button_scan.configure(state='disabled')
    statuslabel.configure(text=" .. reset!")
    colorLabel.configure(bg=root.cget('bg'))
    # reset data
    excelfiles.clear()
    reports.clear()
    global color_search
    color_search = '#FF0000'

# MAIN script
# create root Tkinter element
root = Tk()

excelfiles = list()
reports = list()

global color_search
color_search = '#FF0000'

# define status label for info to user
statuslabel = Label(root, text="waiting for reports...")
statuslabel.grid(row=0, column=0, columnspan=6, padx='5', pady='5', sticky='W')

# define listbox for reports
listbox = Listbox(root, height=15, width=50, selectmode='single')
listbox.grid(row=1, column=0, rowspan=4, columnspan=5, padx='5', pady='5', sticky='NSEW')
listbox.bind("<<ListboxSelect>>", selection_changed)

# buttons for user interaction - BASE actions
button_open = Button(root, text="Choose directory", command=openfolder)
button_open.grid(row=7, column=0, columnspan=2, padx='5', pady='5')
button_quit = Button(root, text="Quit", command=root.quit)
button_quit.grid(row=8, column=4, padx='5', pady='5')
button_reset = Button(root, text="Reset", command=reset)
button_reset.grid(row=7, column=4, padx='5', pady='5')
button_scan = Button(root, text="Start evaluation", command=scan_thread, state='disabled')
button_scan.grid(row=7, column=2, columnspan=2, padx='5', pady='5')
button_color = Button(root, text="Choose color", command=choose_color)
button_color.grid(row=8, column=0, columnspan=2, padx='5', pady='5')

colorLabel = Label(root, text="* current search color *", bg=color_search)
colorLabel.grid(row=8, column=2, columnspan=2, padx='5', pady='5')


if __name__ == '__main__':
    # app settings
    root.title("Excel Report Evaluator")
    # prevent resizing
    root.resizable(False, False)

    root.mainloop()
