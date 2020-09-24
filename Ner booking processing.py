import pandas as pd
import xlsxwriter

from pathlib import Path
import os

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

rh_columns_required = {'Erev RH':{
                            'Men':'I wish to attend the EREV ROSH HASHANA MINCHA & MAARIV minyan (Mens section)',
                            'Women':'I wish to attend the EREV ROSH HASHANA MINCHA & MAARIV minyan (womens section)'
                        },
                        'Day 1 Men':{
                            1:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 1 (mens section)',
                            2:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 1 (mens section).1',
                            3:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 1 (mens section).2'
                        },
                        'Day 1 Women':{
                            1:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 1 (Womens section)',
                            2:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 1 (Womens section).1',
                            3:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 1 (Womens section).2'
                        },
                       'Mincha Day 1':{
                            'Men':'I wish to attend the ROSH HASHANA MINCHA on DAY 1 minyan (Mens section)',
                            'Women':'I wish to attend the ROSH HASHANA MINCHA on Day 1 minyan (womens section)'
                        },
                        'Maariv Day 2':{
                            'Men':'I wish to attend the MAARIV minyan on Rosh Hashanah Day 2 (Mens section)',
                            'Women':'I wish to attend MAARIV minyan on Rosh Hashanah Day 2  (womens section)'
                        },
                        'Day 2 Men':{
                            1:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 2 (mens section)',
                            2:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 2 (mens section).1',
                            3:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 2 (mens section).2'
                        },
                        'Day 2 Women':{
                            1:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 2 (Womens section)',
                            2:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 2 (Womens section).1',
                            3:'I wish to attend the following ROSH HASHANA MORNING minyanim on DAY 2 (Womens section).2'
                        },
                        'End of Day 2':{
                            'Men':'I wish to attend the ROSH HASHANA MINCHA, SHIUR & MAARIV on DAY 2 minyan (Mens section)',
                            'Women':'I wish to attend the ROSH HASHANA MINCHA, SHIUR & MAARIV on Day 2 minyan (womens section)'
                        },
                        'Children':{
                            'Day 1':'I wish to attend a Rosh Hashanah Day 1 CHILDREN service (Ner campus)',
                        }
    }

shabbos_columns_required = {'Men Shacharit':{'':'I wish to attend the following SHABBAT MORNING minyan (mens section)'},
            'Women Shacharit':{'':'I wish to attend the following SHABBAT MORNING minyan (womens section)'},
            'Kabbalat Shabbat': {'Men':'I wish to attend the KABBALAT SHABBAT minyan at the end of the week (mens section)',
                                'Women':'I wish to attend the KABBALAT SHABBAT minyan at the end of the week (womens section)'},
            'Mincha':{'Men':'I wish to attend the SHABBAT MINCHA minyan (mens section)',
                    'Women':'I wish to attend the SHABBAT MINCHA minyan (womens section)'},
            'Children Service':{'':'I wish to attend a shabbat morning CHILDREN service'},
    }

def run_process():
    input_fp = inputs_filepath.get()
    outputs_fp = outputs_filepath.get()
    title = label.get()

    selection_value = selection.get()
    if selection_value == "Shabbos":
        columns_required = shabbos_columns_required
    elif selection_value == "RH":
        columns_required = rh_columns_required

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
            # Get unique list of options within the column
            unique_options = info_item.dropna().unique()
            # If > 1 option, sort by trying to find the time and turning it into a number
            if len(unique_options)!=1:
                UO_sort = sorted(unique_options, key=FindMIfFloat)
            else: # if only 1 option, no sorting needed
                UO_sort = unique_options

            # Loop through the options
            for option in UO_sort:
                # Find which rows match the option
                filterrows = info_item == option
                # Title the column in Excel based on selection above
                if col == '' or isinstance(col, int):
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

custom_JSON = StringVar()
def show_custom(self):
    selection_value = selection.get()
    if selection_value=="Custom":
        global custom_label
        custom_label = ttk.Label(mainframe, text="Custom JSON")
        custom_label.grid(column=1, row=Selection_Row+1, sticky=E)
        global JSON_entry
        JSON_entry = ttk.Entry(mainframe, textvariable=custom_JSON)
        JSON_entry.grid(column=2, row=Selection_Row+1, sticky=(W, E), columnspan=2)
    else:
        custom_label.grid_forget()
        JSON_entry.grid_forget()



ttk.Label(mainframe, text="Select type").grid(column=1, row=Selection_Row, sticky=E)
selection = StringVar()
#Options = ["Shabbos", "RH", "Custom"]
Options = ["Shabbos", "RH"]
#selection_dropdown = ttk.OptionMenu(mainframe, selection, "Pick", *Options, command=show_custom)
selection_dropdown = ttk.OptionMenu(mainframe, selection, "Shabbos", *Options)
selection_dropdown.grid(column=2, row=Selection_Row, sticky=(W))


ttk.Button(mainframe, text="Run", command=run_process).grid(column=4, row=5, sticky=E)


for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.bind('<Return>', run_process)
root.mainloop()
