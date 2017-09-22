import os
import json
import pyperclip
import tkinter as tk
from tkinter import messagebox
import win32com.client
import csv
from PIL import ImageGrab

VP_JSON = r'values_pairs.json'
VP_CSV = r'H:Sales\Public\National Accounts\National Accounts\jt\value_pairs.csv'


def create_lookup():
    """
    Creates the local values_pairs.json file that is used for matching.
    """

    with open(VP_CSV) as f:
        values_pairs = dict(csv.reader(f, delimiter=','))

    with open(VP_JSON, 'w') as f_obj:
        json.dump(values_pairs, f_obj)


def match_pairs():
    """Loads local json file into a dict, and matches based on input."""

    with open(VP_JSON) as f_obj:
        values_pairs = json.load(f_obj)

    output_list = []
    errors = []
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
            if clean_lookup in values_pairs.keys():
                isbn = list(values_pairs.keys())[
                        list(values_pairs.values()).index(clean_lookup)]
                output_list.append(isbn)
            else:
                errors.append(clean_lookup)
                output_list.append("Not found")

    if len(errors):
        error_msg = '\n'.join(errors)

        if len(errors) == 1:
            msg = "A valid value for {} could not be found.".format(error_msg)
        else:
            msg = ("Valid values for the following could not be found: \n\n"
                   .format(error_msg)
                   )
        hide_root()
        messagebox.showinfo("No Match Found", msg)

    



def show_unmatched():
    """Informs user of no match, and allows them to see what they input."""

    # Limits displayed input to 30 characters for readability.
    if len(lookup) > 30:
        lookup_error = "{} . . . ".format(lookup[:30])
    else:
        lookup_error = lookup

    no_match_msg = "There was no valid match for:\n\n {}".format(lookup_error)
    hide_root()
    messagebox.showinfo("No Match Found", no_match_msg)


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


def hide_root():
    """Just hides the Tkinter terminal window so it doesn't pop up."""
    root = tk.Tk()
    root.withdraw()


def email_unexpected_error():
    """Tells user an unexpected error occurred, and encourages emailing JT."""

    hide_root()

    error_msg = ("There was an unexpected error. Would you like to email JT?\n" 
                 "Doing so will help make improvements so that this doesn't "
                 "happen again.")

    if messagebox.askyesno("Unexpected Enabler Error", error_msg):

        o = win32com.client.Dispatch("Outlook.Application")

        msg = o.CreateItem(0)
        msg.Subject = "Enabler Error Submission"
        msg.HTMLBody = (
            "Hi JT,<br><br>I encountered an error while using Enabler. "
            "Here is what I was doing when the error occurred:<br><br>"
            "<b>What I was trying to convert: </b> {}"
            "<br><b>More details about what I was doing: (As detailed as "
            "possible, please)</b> "
            "<br><br>Eternally grateful,<br>".format(lookup))

        msg.To = "jt@workman.com"

        msg.Display()


###############################################################################
lookup = pyperclip.paste().splitlines()

while True:

    # if not ImageGrab.grabclipboard():
    #
    #
    # else:
    #     hide_root()
    #     messagebox.showinfo("Charon Error", "Your clipboard contains an image. "
    #                                     "Please search for text only.")
    #     break

    try:
        match_pairs()
        break

    except FileNotFoundError:
        create_lookup()

    except (KeyError, ValueError):
        if update_available():
            if should_update_json():
                create_lookup()
            else:
                break
        else:
            show_unmatched()
            break

    #except Exception:
        #email_unexpected_error()
        #break
