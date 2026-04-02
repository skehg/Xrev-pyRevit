# -*- coding: utf-8 -*-
from pyrevit import revit, forms
from Autodesk.Revit.DB import SaveAsOptions
from Autodesk.Revit.UI import (
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogCommandLinkId,
    TaskDialogResult,
)
from datetime import datetime
import os
import re
import shutil
import sys

doc = revit.doc


def exit_with_message(message, title):
    forms.alert(message, title=title)
    sys.exit()


def ask_original_handling():
    dialog = TaskDialog("Original Family Handling")
    dialog.MainInstruction = "Choose what to do with the original family file"
    dialog.MainContent = (
        "This tool will create a compacted family by saving to a temporary "
        "'-new#' file and then saving back to the base file name."
    )
    dialog.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink1,
        "Move original to Superseded folder with timestamp"
    )
    dialog.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink2,
        "Delete original family file"
    )
    dialog.CommonButtons = TaskDialogCommonButtons.Cancel

    result = dialog.Show()
    if result == TaskDialogResult.CommandLink1:
        return "move"
    if result == TaskDialogResult.CommandLink2:
        confirm = TaskDialog("Confirm Delete")
        confirm.MainInstruction = "Delete the original family file?"
        confirm.MainContent = "This cannot be undone."
        confirm.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
        return "delete" if confirm.Show() == TaskDialogResult.Yes else None
    return None


def find_next_new_path(folder, base_name, extension):
    counter = 1
    while counter <= 10000:
        candidate_name = "{}-new{}{}".format(base_name, counter, extension)
        candidate_path = os.path.join(folder, candidate_name)
        if not os.path.exists(candidate_path):
            return candidate_path
        counter += 1
    return None


def get_timestamped_path(superseded_folder, file_name, extension):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_candidate = "{}_{}{}".format(file_name, timestamp, extension)
    candidate_path = os.path.join(superseded_folder, base_candidate)
    index = 1
    while os.path.exists(candidate_path):
        fallback_name = "{}_{}_{}{}".format(file_name, timestamp, index, extension)
        candidate_path = os.path.join(superseded_folder, fallback_name)
        index += 1
    return candidate_path

# ---------------------------------------------------------
# Validate Document is a Family
# ---------------------------------------------------------
if not doc.IsFamilyDocument:
    exit_with_message(
        "This tool only works in the Family Editor.\n"
        "Please open a family document (.rfa) to use this feature.",
        "Not a Family Document"
    )

# ---------------------------------------------------------
# Get Original File Info
# ---------------------------------------------------------
original_path = doc.PathName

if not original_path or original_path == "":
    exit_with_message(
        "The family document must be saved before using this tool.",
        "Unsaved Document"
    )

# Parse file path
directory = os.path.dirname(original_path)
original_filename = os.path.basename(original_path)
name_without_ext = os.path.splitext(original_filename)[0]
file_extension = os.path.splitext(original_filename)[1]

# ---------------------------------------------------------
# Remove existing "-new#" suffix if present to get true base name
# ---------------------------------------------------------
# Pattern: matches "-new" followed by digits at the end
base_name = re.sub(r'-new\d+$', '', name_without_ext)
final_filename = "{}{}".format(base_name, file_extension)
final_path = os.path.join(directory, final_filename)

action = ask_original_handling()
if not action:
    sys.exit()

# ---------------------------------------------------------
new_filename = find_next_new_path(directory, base_name, file_extension)
if not new_filename:
    exit_with_message(
        "Too many saved versions found. Please clean up old files.",
        "Error"
    )

txt_path = os.path.join(directory, "{}.txt".format(base_name))
txt_orig_path = "{}.orig".format(txt_path)
txt_was_renamed = False

try:
    if os.path.exists(txt_path):
        if os.path.exists(txt_orig_path):
            exit_with_message(
                "A temporary '.orig' txt file already exists:\n{}\n"
                "Resolve it before running this tool.".format(txt_orig_path),
                "TXT Rename Conflict"
            )
        os.rename(txt_path, txt_orig_path)
        txt_was_renamed = True

    save_options = SaveAsOptions()
    save_options.OverwriteExistingFile = False
    save_options.Compact = True

    # Step 1: SaveAs to temporary -new# file.
    doc.SaveAs(new_filename, save_options)

    # Step 2: Move or delete the original file.
    if action == "move":
        superseded_folder = os.path.join(directory, "Superseded")
        if not os.path.isdir(superseded_folder):
            os.makedirs(superseded_folder)
        archive_path = get_timestamped_path(superseded_folder, name_without_ext, file_extension)
        shutil.move(original_path, archive_path)
    elif action == "delete":
        os.remove(original_path)

    # Step 3: SaveAs back to base/original family file name.
    if os.path.exists(final_path):
        exit_with_message(
            "Cannot save back to original name because file exists:\n{}".format(final_path),
            "Final Save Conflict"
        )
    doc.SaveAs(final_path, save_options)

    # Step 4: Delete temporary -new# file.
    if os.path.exists(new_filename):
        os.remove(new_filename)

except Exception as error:
    exit_with_message(
        "An error occurred while compact-saving the family:\n\n{}".format(str(error)),
        "Save Error"
    )
finally:
    if txt_was_renamed and os.path.exists(txt_orig_path):
        if os.path.exists(txt_path):
            os.remove(txt_path)
        os.rename(txt_orig_path, txt_path)

forms.alert(
    "Family compact-save complete!\n\n"
    "Current file: {}".format(os.path.basename(final_path)),
    title="Save Successful"
)
