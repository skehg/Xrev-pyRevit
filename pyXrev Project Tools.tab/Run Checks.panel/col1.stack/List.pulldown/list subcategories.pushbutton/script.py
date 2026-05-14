# -*- coding: utf-8 -*-
"""
Export Subcategories to CSV

This pyRevit script lists all subcategories in the current Revit project,
including relevant identifiers, and exports the data to a CSV file.

Exported fields:
- Subcategory Name
- Subcategory Id
- Parent Category Name
- Parent Category Id
- BuiltInCategory
- Is Built-In Subcategory (Yes/No)

Compatible with Revit 2021–2025 and pyRevit 4.8.13+ / 5.x (IronPython 2.7).
"""

__title__ = "Export Subcategories to CSV"
__author__ = "Your Name"
__doc__ = "Exports all subcategories in the current project to a CSV file."

# --- Imports ---
from pyrevit import revit, script, forms
import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Category, BuiltInCategory

# --- Helper Functions ---

def get_builtincategory_str(category):
    try:
        bic = getattr(category, 'BuiltInCategory', None)
        if bic is not None:
            return str(bic)
    except Exception:
        pass
    return "Invalid"

def is_builtin_from_bic_string(bic_str):
    """
    Determines if the BuiltInCategory string is valid.
    """
    return bic_str != "Invalid"

def safe_getattr(obj, attr, default=""):
    try:
        return getattr(obj, attr, default)
    except Exception:
        return default

# --- Main Logic ---

doc = revit.doc

# Prepare CSV header
csv_data = [
    [
        "Subcategory Name",
        "Subcategory Id",
        "Parent Category Name",
        "Parent Category Id",
        "BuiltInCategory",
        "Is Built-In Subcategory"
    ]
]

# Iterate all categories and extract subcategories
for category in doc.Settings.Categories:
    subcats = safe_getattr(category, "SubCategories", None)
    if subcats and subcats.Size > 0:
        for subcat in subcats:
            subcat_name = subcat.Name
            subcat_id = subcat.Id.IntegerValue
            parent_name = category.Name
            parent_id = category.Id.IntegerValue
            bic_str = get_builtincategory_str(subcat)
            is_builtin = "Yes" if is_builtin_from_bic_string(bic_str) else "No"
            csv_data.append([
                subcat_name,
                subcat_id,
                parent_name,
                parent_id,
                bic_str,
                is_builtin
            ])

# --- Prompt User for Save Location ---

csv_path = forms.save_file(
    file_ext='csv',
    title='Save Subcategories CSV'
)

if not csv_path:
    forms.alert("Export cancelled. No file was saved.", exitscript=True)

# --- Write CSV File ---

try:
    script.dump_csv(csv_data, csv_path)
    forms.alert(
        "Export successful!\n\nFile saved to:\n{}".format(csv_path),
        title="Export Complete"
    )
    script.show_file_in_explorer(csv_path)
except Exception as e:
    forms.alert(
        "An error occurred while exporting the CSV:\n\n{}".format(str(e)),
        exitscript=True
    )