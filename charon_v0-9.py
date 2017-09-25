import os
import json
import pyperclip
import tkinter as tk
from tkinter import messagebox
import win32com.client
import csv

"""Charon v0.9"""

VP_JSON = r'values_pairs.json'
VP_CSV = r'H:\Sales\Public\National Accounts\National Accounts\jt\value_pairs.csv'


def create_lookup():
    """
    Creates the local values_pairs.json file that is used for matching.
    """

    with open(VP_CSV) as f:
        values_pairs = dict(csv.reader(f, delimiter=','))

    with open(VP_JSON, 'w') as f_obj:
        json.dump(values_pairs, f_obj)


def hide_tk():
    """Just hides the Tkinter terminal window so it doesn't pop up."""
    root = tk.Tk()
    root.withdraw()


def update_available():
    """Checks to see if server's master CSV file is newer than local JSON."""
    csv_time = os.path.getmtime(VP_CSV)
    json_time = os.path.getmtime(VP_JSON)

    if csv_time > json_time:
        return True


def should_update_json():
    """If no match might be due to outdated JSON, allows user to update."""
    hide_tk()

    if messagebox.askyesno("Update?",
                           "At least one value could not be matched.\n"
                           "\nWould you like to update and try again?\n"
                           "(If you update and do not see an error, a match "
                           "was found.)"
                           ):
        return True


def show_unmatched():
    """Informs user of no match, and allows them to see what they input."""

    error_msg = '\n'.join(errors)

    if len(errors) == 1:
        msg = "A valid value for {} could not be found.".format(error_msg)

    else:
        msg = ("Valid values for the following could not be found "
               "({} out of {}): \n\n{}".format(len(errors), len(lookup),
                                               error_msg)
               )
    hide_tk()
    messagebox.showinfo("No Match Found", msg)


def email_unexpected_error():
    """Tells user an unexpected error occurred, and encourages emailing JT."""

    hide_tk()

    error_msg = ("There was an unexpected error. Would you like to email JT?\n" 
                 "Doing so will help make improvements so that this doesn't "
                 "happen again.")

    if messagebox.askyesno("Unexpected Enabler Error", error_msg):

        o = win32com.client.Dispatch("Outlook.Application")

        msg = o.CreateItem(0)
        msg.Subject = "Charon Error Submission"
        msg.HTMLBody = (
            "Hi JT,<br><br>I encountered an error while using Charon. "
            "Here is what I was doing when the error occurred:<br><br>"
            "<b>What I was trying to convert: </b> {}"
            "<br><b>More details about what I was doing: (As detailed as "
            "possible, please)</b> "
            "<br><br>Eternally grateful,<br>".format("\n".join(lookup)))

        msg.To = "jt@workman.com"

        msg.Display()


###############################################################################

lookup = pyperclip.paste().splitlines()
multiple_columns = False

try:
    while True:
        output_list = []
        errors = []
        try:
            with open(VP_JSON) as f_obj:
                values_pairs = json.load(f_obj)
        except FileNotFoundError:
            create_lookup()
            continue

        for value in lookup:
            clean_lookup = value.strip().replace("-", "")

            # Looks for corresponding Product ID of the ISBN.
            if len(clean_lookup) == 13 and clean_lookup.startswith("978"):
                if clean_lookup in values_pairs:
                    output_list.append(values_pairs[clean_lookup])
                else:
                    errors.append(clean_lookup)
                    output_list.append("Not found")

            # Maintains blank rows in output.
            elif clean_lookup == "":
                output_list.append("")

            # Checks to see if the person accidentally copied multiple columns.
            elif "\t" in clean_lookup:
                multiple_columns = True
                break

            # Looks for corresponding ISBN of the Product ID.
            else:
                if clean_lookup in values_pairs.values():
                    isbn = list(values_pairs.keys())[list(
                        values_pairs.values()).index(clean_lookup)]
                    output_list.append(isbn)
                else:
                    errors.append(clean_lookup)
                    output_list.append("Not found")

        if len(errors):
            if update_available():
                if should_update_json():
                    create_lookup()
                    continue

            show_unmatched()

        if output_list and not multiple_columns:
            copy_output = "\n".join(output_list)
            pyperclip.copy(copy_output)

            break

        elif not output_list and not multiple_columns:
            hide_tk()
            messagebox.showinfo("Invalid Input", "It looks like you have no "
                                "text in your clipboard. Please double check "
                                "and try again.")
            break

        elif multiple_columns:
            hide_tk()
            messagebox.showinfo("Invalid Input", "It looks like you've copied "
                                "multiple columns, which is currently "
                                "unsupported. Please select only one column to"
                                " convert.")
            break

except:
    email_unexpected_error()
