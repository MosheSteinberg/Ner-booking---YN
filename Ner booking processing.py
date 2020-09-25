import pandas as pd
import xlsxwriter

from pathlib import Path
import os, sys, json

from tkinter import ttk
from tkinter import Tk, StringVar, N, W, E, S, IntVar
from tkinter import filedialog, PhotoImage
from tkcalendar import DateEntry
from datetime import date, timedelta
from pandas import to_datetime

def FindMIfFloat(x):
    try:
        return float(x[:(x.find('m ')-1)])
    except ValueError:
        return x

def IsInteger(x):
    try:
        int(x)
        return True
    except:
        return False

def run_process():
    input_fp = inputs_filepath.get()
    outputs_fp = outputs_filepath.get()
    title = label.get()

    selection_value = selection.get()
    with open(selection_value, mode='r') as ner_file:
        json_value = ner_file.read()
        columns_required = json.loads(json_value)

    raw_data = pd.read_csv(input_fp, error_bad_lines=False)
    
    firstname_column = 7
    surname_column = 8
    raw_data.columns.values[firstname_column] = 'firstname'
    raw_data.columns.values[surname_column] = 'surname'
    check_deletion = delete_flag.get()

    if check_deletion == 1:
        delete_before_date = pd.to_datetime(delete_before_entry.get_date())
        formatted_date_column = pd.to_datetime(raw_data['Submission Date'], format='%d/%m/%y %H:%M:%S')
        raw_data = raw_data[formatted_date_column > delete_before_date]

    writer = pd.ExcelWriter(outputs_fp, engine='xlsxwriter')
    workbook = writer.book
    format_cells = workbook.add_format({'font_size':22})
    format_titles = workbook.add_format()
    format_titles.set_text_wrap()
    format_titles.set_bold()
    format_titles.set_border()
    format_titles.set_align('center')
    format_titles.set_align('vcenter')

    # Loop through sheets
    for item, column_names in columns_required.items():
        # Get the name for the sheet
        sheet_name = str(item)
        ## Start the column count at 0
        number = 0
        ## Loop through the columns
        for col, column_name in column_names.items():
            # Pick out the Info
            info_item = raw_data[column_name]

            info_item_save_commas = info_item.str.replace(', ', '>>')

            split_column = info_item_save_commas.str.get_dummies(sep=',')
            # Get unique list of options within the column
            unique_options = [val.replace('>>', ', ') for val in list(split_column.columns.values)]
            print(unique_options)
            # If > 1 option, sort by trying to find the time and turning it into a number
            UO_sort = sorted(unique_options, key=FindMIfFloat)
            # Loop through the options
            for option in UO_sort:
                # Find which rows match the option
                filterrows = split_column[option.replace(', ', '>>')]==1
                # Title the column in Excel based on selection above
                if col == '' or IsInteger(col):
                    name = option
                else:
                    name = col
                # Filter the attendees and apply a header to the column
                ListOfAttendees_Unordered = (raw_data['surname'] + ', ' + raw_data['firstname'])[filterrows].rename(name)
                # Order the attendees alphabetically
                ListOfAttendees_Ordered = ListOfAttendees_Unordered.sort_values()
                # Write the list to Excel in the 3rd row
                ListOfAttendees_Ordered.to_excel(writer, sheet_name, index=False, startcol=number, startrow=2)
                # Select the column and set its width
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column(number, number, len(name)+10)
                worksheet.write(2, number, name, format_titles)
                #print(number)
                # Write the name in the top left
                worksheet.write('A1', sheet_name, format_cells)
                # Write the title in next row
                worksheet.write('A2', title, format_cells)
                # Write count next to table
                worksheet.write(2, number + 1, len(ListOfAttendees_Ordered))
                # Set height of first row
                worksheet.set_row(0, 22)
                worksheet.set_row(2, 30)
                worksheet.fit_to_pages(1,1)
                # Increment 2 columns across for the next list
                number += 2
                
    # Save the workbook
    writer.save()

    # Open the workbook for the user
    os.startfile(outputs_fp)

    return 0

def file_explore_inputs():
    filename = filedialog.askopenfilename(initialdir=os.path.join(Path.home(), 'Downloads'),
                                          title='Select a file',
                                          filetypes=(('CSV files', '*.csv*'),))
    inputs_filepath.set(filename)

def file_explore_outputs():
    filename = filedialog.asksaveasfilename(initialdir=os.path.join(Path.home(), 'Documents'),
                                          title='Select a file',
                                          filetypes=(('Excel files', '*.xlsx'),))
    if filename[-5:] != '.xlsx':
        filename = filename + '.xlsx'
    outputs_filepath.set(filename)

root = Tk()
root.title("Ner booking process")
root.resizable(0,0)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)	

inputs_filepath = StringVar()
outputs_filepath = StringVar()
label = StringVar()
delete_before = StringVar()

input_row = 1
output_row = 2
date_row = 3
Label_Row = 4
Selection_Row = 5

ttk.Label(mainframe, text="Location of csv file").grid(column=1, row=input_row, sticky=W)
inputs_filepath_entry = ttk.Entry(mainframe, width=60, textvariable=inputs_filepath)
inputs_filepath_entry.grid(column=2, row=input_row, sticky=(W, E), columnspan=2)
ttk.Button(mainframe, text="Browse", command=file_explore_inputs).grid(column=4, row=input_row, sticky=W)

ttk.Label(mainframe, text="Name of output file").grid(column=1, row=output_row, sticky=W)
outputs_filepath_entry = ttk.Entry(mainframe, textvariable=outputs_filepath)
outputs_filepath_entry.grid(column=2, row=output_row, sticky=(W, E), columnspan=2)
ttk.Button(mainframe, text="Browse", command=file_explore_outputs).grid(column=4, row=output_row, sticky=W)

ttk.Label(mainframe, text="Delete entries from before:").grid(column=1, row=date_row, sticky=W)
default_date = date.today() - timedelta(days=1)
delete_before_entry = DateEntry(mainframe, locale='en_UK')
delete_before_entry.set_date(default_date)
delete_before_entry.grid(column=3, row=date_row, sticky=(W, E))
delete_before_entry.grid_remove()

def show_or_hide_date():
    check_flag = delete_flag.get()
    if check_flag==0:
        delete_before_entry.grid_remove()
    elif check_flag==1:
        delete_before_entry.grid(column=3, row=date_row, sticky=(W, E))
delete_flag = IntVar()
delete_flag.set(1)
date_toggle = ttk.Checkbutton(mainframe, variable=delete_flag, command=show_or_hide_date)
date_toggle.grid(column=2, row=date_row, sticky=(W, E))

label = StringVar()
ttk.Label(mainframe, text="Label").grid(column=1, row=Label_Row, sticky=W)
label_entry = ttk.Entry(mainframe, textvariable=label)
label_entry.grid(column=2, row=Label_Row, columnspan=2, sticky=(W, E))

ttk.Label(mainframe, text="Select type").grid(column=1, row=Selection_Row, sticky=E)
selection = StringVar()

if getattr(sys, 'frozen', False):
    current_directory = os.path.dirname(sys.executable)
else:
    current_directory = os.getcwd()
Options = [json_file for json_file in os.listdir(current_directory) if json_file.endswith('.ner')]
if 'Shabbos.ner' in Options:
    Default = 'Shabbos.ner'
else:
    Default = Options[0]

selection_dropdown = ttk.OptionMenu(mainframe, selection, Default, *Options)
selection_dropdown.grid(column=2, row=Selection_Row, sticky=(W))


ttk.Button(mainframe, text="Run", command=run_process).grid(column=4, row=5, sticky=E)


for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.bind('<Return>', run_process)
root.mainloop()
