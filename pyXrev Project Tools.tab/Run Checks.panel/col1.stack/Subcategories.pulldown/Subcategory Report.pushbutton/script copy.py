# -*- coding: utf-8 -*-
# Audit Family Subcategories by Parent Category
# CSV filename is prefixed with the selected parent category (sanitized)
# Only writes rows where subcat.BuiltInCategory is "INVALID"

import csv
import os
import codecs
import re
import subprocess
from Autodesk.Revit.DB import FilteredElementCollector, Family, CategoryType
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script
from pyrevit.forms import SelectFromList, pick_folder

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# Prompt user to browse for a folder
csv_folder = pick_folder(title="Select folder to save FamilySubcategoriesReport")
if not csv_folder:
    script.exit()

# Get parent categories (model categories with no parent)
all_categories = doc.Settings.Categories
parent_categories = sorted(
    [cat.Name for cat in all_categories if cat.CategoryType == CategoryType.Model and cat.Parent is None]
)

# Prompt user
selected = SelectFromList.show(parent_categories, title="Select Parent Category", multiselect=False)
if not selected:
    TaskDialog.Show("Cancelled", "No category selected. Script aborted.")
    script.exit()

# Normalize SelectFromList return type to a single string
selected_cat_name = selected[0] if isinstance(selected, (list, tuple)) else selected
selected_cat = next((cat for cat in all_categories if cat.Name == selected_cat_name), None)
if not selected_cat:
    TaskDialog.Show("Error", "Could not resolve selected category.")
    script.exit()

# Sanitize the selected category name for safe filename use
safe_name = re.sub(r'[<>:"/\\|?*\n\r\t]+', '_', selected_cat_name).strip()
safe_name = safe_name.replace(' ', '_')
if not safe_name:
    safe_name = "Category"

csv_filename = "{}_FamilySubcategoriesReport.csv".format(safe_name)
csv_path = os.path.join(csv_folder, csv_filename)

selected_cat_id = selected_cat.Id
report_data = []

# Iterate families directly from the collector
fam_collector = FilteredElementCollector(doc).OfClass(Family)
for fam in fam_collector:
    # Skip families without a category or whose family category doesn't match the selected parent
    try:
        fam_cat = fam.FamilyCategory
    except Exception:
        fam_cat = None

    if not fam_cat or fam_cat.Id != selected_cat_id:
        continue

    fam_doc = None
    try:
        fam_doc = doc.EditFamily(fam)  # open family document once
        root_cat = fam_doc.OwnerFamily.FamilyCategory

        # Ensure it still matches the selected parent
        if not root_cat or root_cat.Id != selected_cat_id:
            continue

        for subcat in root_cat.SubCategories:
            # Determine BuiltInCategory value; treat exceptions or invalid IDs as "INVALID"
            try:
                bic = subcat.BuiltInCategory
                bic_name = str(bic) if bic != -1 else "INVALID"
            except Exception:
                bic_name = "INVALID"

            # Only include rows where BuiltInCategory is "INVALID"
            if bic_name == "INVALID":
                report_data.append((fam.Name, subcat.Name, bic_name))

    except Exception:
        # If a family fails to open, skip it (we only write rows where subcat.BuiltInCategory == "INVALID")
        pass
    finally:
        try:
            if fam_doc:
                fam_doc.Close(False)
        except Exception:
            pass

# Write CSV with UTF-8 BOM so Excel opens it correctly
with codecs.open(csv_path, "w", encoding="mbcs") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Family Name", "Subcategory Name", "BuiltInCategory", "New Subcategory"])
    for row in report_data:
        writer.writerow(row)

output.print_md("✅ CSV report created for category **{}**: `{}`\n\nRows included: {}".format(
    selected_cat_name, csv_path, len(report_data)
))

# Attempt to open the file for the user (Windows/Revit environment)
opened = False
try:
    os.startfile(csv_path)
    opened = True
except Exception:
    try:
        subprocess.Popen(['explorer', csv_path])
        opened = True
    except Exception:
        try:
            script.open_url("file:///" + csv_path.replace("\\", "/"))
            opened = True
        except Exception:
            opened = False

if not opened:
    TaskDialog.Show("Notice", "CSV created but could not be opened automatically.\nFile saved to:\n{}".format(csv_path))