import pandas as pd
import xlsxwriter

from pathlib import Path
import os

from tkinter import ttk
from tkinter import Tk, StringVar, N, W, E, S
from tkinter import filedialog

def FindMIfFloat(x):
    try:
        return float(x[:(x.find('m ')-1)])
    except ValueError:
        return x

def run_process():
    input_fp = inputs_filepath.get()
    outputs_fp = outputs_filepath.get()

    raw_data = pd.read_csv(input_fp, error_bad_lines=False)
    firstname_column = 7
    surname_column = 8
    raw_data.columns.values[firstname_column] = 'firstname'
    raw_data.columns.values[surname_column] = 'surname'

    writer = pd.ExcelWriter(outputs_fp, engine='xlsxwriter')
    workbook = writer.book
    format_titles = workbook.add_format({'text_wrap':True})

    columns_required = {'Erev RH':{
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

       ## Loop through sheets
    for item in columns_required:
        # Get the name of the columns in the csv for this sheet
        column_names = columns_required[item]
        # Get the name for the sheet
        sheet_name = str(item)
        ## Start the column count at 0
        number = 0
        ## Loop through the columns
        for col in column_names:
            # Pick out the column name
            column_name = column_names[col]
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
                #print(number)
                # Increment 2 columns across for the next list
                number += 2
                # Write the name in the top left
                worksheet.write('A1', sheet_name)
                # Set height of first row
                worksheet.set_row(0, 22)
                worksheet.set_row(2, None, format_titles)

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

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)	

inputs_filepath = StringVar()
outputs_filepath = StringVar()
label = StringVar()
delete_before = StringVar()

ttk.Label(mainframe, text="Location of csv file").grid(column=1, row=1, sticky=W)
inputs_filepath_entry = ttk.Entry(mainframe, width=60, textvariable=inputs_filepath)
inputs_filepath_entry.grid(column=2, row=1, sticky=(W, E))
ttk.Button(mainframe, text="Browse", command=file_explore_inputs).grid(column=3, row=1, sticky=W)

ttk.Label(mainframe, text="Name of output file").grid(column=1, row=2, sticky=W)
outputs_filepath_entry = ttk.Entry(mainframe, textvariable=outputs_filepath)
outputs_filepath_entry.grid(column=2, row=2, sticky=(W, E))
ttk.Button(mainframe, text="Browse", command=file_explore_outputs).grid(column=3, row=2, sticky=W)

ttk.Button(mainframe, text="Run", command=run_process).grid(column=3, row=5, sticky=W)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)


root.bind('<Return>', run_process)
root.mainloop()
