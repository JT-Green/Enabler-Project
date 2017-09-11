import os
import json
import pyperclip
import openpyxl
import tkinter as tk
from tkinter import messagebox
import win32com.client
import csv

VP_JSON = r'values_pairs.json'
VP_XLSX = r'G:\jt\prod_id.xlsx'
VP_CSV = r'G:\jt\value_pairs.csv'


def create_lookup():
    """Creates the values_pairs.json file"""

    with open(VP_CSV) as f:
        f.readline()  # ignore first line (header)
        values_pairs = dict(csv.reader(f, delimiter=','))


    # wb = openpyxl.load_workbook(VP_XLSX)
    # sheet = wb.active

    # for row in range(2, sheet.max_row + 1):
    #     lookup_input = sheet['A' + str(row)].value
    #     lookup_output = sheet['B' + str(row)].value
    #
    #     values_pairs[lookup_input] = lookup_output

    with open(VP_JSON, 'w') as f_obj:
        json.dump(values_pairs, f_obj)


def match_pairs():
    """ DOCUMENTATION TK"""
    with open(VP_JSON) as f_obj:
        values = json.load(f_obj)

    if len(lookup) == 13 and lookup[0:3] == "978":
        corresponding = values[lookup]
        pyperclip.copy(corresponding)

    elif len(lookup) <= 8 and lookup[0].isdigit():
        corresponding = list(values.keys())[list(values.values()).index(lookup)]
        pyperclip.copy(corresponding)

    else:
        raise KeyError

def show_unmatched():
    """ DOCUMENTATION TK"""
    if len(lookup) > 30:
        lookup_error = lookup[:30] + " . . ."
    else:
        lookup_error = lookup

    error_msg = "There was no valid match for :\n\n" + lookup_error
    hide_root()
    messagebox.showinfo("Enabler Error", error_msg)


def update_available():
    """ DOCUMENTATION TK"""
    csv_time = os.path.getmtime(VP_CSV)
    json_time = os.path.getmtime(VP_JSON)
    # xlsx_time = os.path.getmtime(VP_XLSX)

    if csv_time > json_time:
        return True


def should_update_json():
    """ DOCUMENTATION TK"""
    hide_root()

    if messagebox.askyesno("Update?",
                "There was no valid match."
                + "\nThere is an update available for the values list."
                + "\nWould you like to update and try again?"
        ):
        return True


def hide_root():
    """ Just hides the Tkinter terminal window so it doesn't pop up. """
    root = tk.Tk()
    root.withdraw()


def email_unexpected_error():
    """ Tells user an unexpected error occurred, and allows emailing JT. """

    o = win32com.client.Dispatch("Outlook.Application")

    msg = o.CreateItem(0)
    msg.Subject = "Enabler Error Submission"
    msg.HTMLBody = (
        "Hi JT,<br><br>I encountered an error while using Enabler." +
        "Here is what I was doing when the error occurred:<br><br>" +
        "[type error conditions here]<br><br>Eternally grateful,<br>")

    msg.To = "jt@workman.com"

    msg.Display()


###############################################################################

lookup = pyperclip.paste().strip()

while True:
    try:
        match_pairs()
        break

    except FileNotFoundError:
        create_lookup()

    except KeyError:
        if update_available():
            if should_update_json():
                create_lookup()
        else:
            show_unmatched()
            break

    except Exception:
        email_unexpected_error()
        break
