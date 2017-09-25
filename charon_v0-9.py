import os
import json
import pyperclip
import tkinter as tk
from tkinter import messagebox
import win32com.client
import csv

VP_JSON = r'values_pairs.json'
VP_CSV = r'C:\Users\JT\Desktop\test_csv.csv'

def create_lookup():
    """
    Creates the local values_pairs.json file that is used for matching.
    """

    with open(VP_CSV) as f:
        values_pairs = dict(csv.reader(f, delimiter=','))

    with open(VP_JSON, 'w') as f_obj:
        json.dump(values_pairs, f_obj)


def hide_root():
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
    hide_root()

    if messagebox.askyesno("Update?",
                           "There was no valid match."
                           "\nThere is an update available for the values list."
                           "\nWould you like to update and try again?"
                           "\n(If you update and do not see an error, a match was found.)"
                           ):
        return True

def show_unmatched():
    """Informs user of no match, and allows them to see what they input."""

    error_msg = '\n'.join(errors)

    if len(errors) == 1:
        msg = "A valid value for {} could not be found.".format(error_msg)

    else:
        msg = ("Valid values for the following could not be found: \n\n"
               .format(error_msg)
               )

    hide_root()
    messagebox.showinfo("No Match Found", msg)

lookup = pyperclip.paste().splitlines()
output_list = []
errors = []

while True:

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

        # Looks for corresponding ISBN of the Product ID.
        else:
            if clean_lookup in values_pairs.values():
                isbn = list(values_pairs.keys())[list(values_pairs.values()).index(clean_lookup)]
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

    copy_output = "\n".join(output_list)
    pyperclip.copy(copy_output)
    break

print(len(output_list))

if not output_list: print("nah")

# If clean_lookup is blank, return blank
# if NOT output_list, error message
