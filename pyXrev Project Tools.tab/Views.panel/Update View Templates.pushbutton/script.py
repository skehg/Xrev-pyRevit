# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, Transaction
)
from pyrevit import forms
import os, csv

# Get active document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# -------------------------------
# 1. Collect all view templates
# -------------------------------
templates = FilteredElementCollector(doc).OfClass(View).ToElements()
templates = [v for v in templates if v.IsTemplate]

if not templates:
    forms.alert("No view templates found.", exitscript=True)

# -------------------------------
# 2. Select templates
# -------------------------------
selected_templates = forms.SelectFromList.show(
    [v.Name for v in templates],
    multiselect=True,
    title="Select View Templates to Update"
)
if not selected_templates:
    forms.alert("No templates selected.", exitscript=True)

templates_to_update = [v for v in templates if v.Name in selected_templates]

# -------------------------------
# 3. Select parent categories (including annotation)
# -------------------------------
all_cats = doc.Settings.Categories
parent_cat_names = [c.Name for c in all_cats if c.Parent is None]

selected_parents = forms.SelectFromList.show(
    parent_cat_names,
    multiselect=True,
    title="Select Parent Categories (Model + Annotation)"
)
if not selected_parents:
    forms.alert("No parent categories selected.", exitscript=True)

categories_to_update = [c for c in all_cats if c.Name in selected_parents]

# -------------------------------
# 4. ON/OFF control
# -------------------------------
visibility_choice = forms.alert(
    "Do you want to turn selected categories ON or OFF?",
    options=["ON", "OFF"]
)
set_hidden = True if visibility_choice == "OFF" else False

# -------------------------------
# 5. Apply changes safely
# -------------------------------
errors = []
t = Transaction(doc, "Update Parent Category Visibility")
try:
    t.Start()
    for vt in templates_to_update:
        for cat in categories_to_update:
            try:
                vt.SetCategoryHidden(cat.Id, set_hidden)
            except Exception as e:
                errors.append("Category '{}' on template '{}': {}".format(cat.Name, vt.Name, e))
    t.Commit()
except Exception as e:
    errors.append("Transaction failed: {}".format(e))
    try:
        t.RollBack()
    except:
        pass
finally:
    try:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
    except:
        pass

# -------------------------------
# 6. User-specified CSV path
# -------------------------------
log_path = forms.save_file(file_ext='csv', title='Save Audit Log As')
if not log_path:
    log_path = os.path.join(os.getenv("USERPROFILE"), "Desktop", "ParentCategoryAudit.csv")

# -------------------------------
# 7. Write audit log
# -------------------------------
try:
    with open(log_path, "w") as f:
        writer = csv.writer(f)
        writer.writerow(["Template", "Parent Category", "Visibility"])
        for vt in templates_to_update:
            for cat in categories_to_update:
                writer.writerow([
                    vt.Name,
                    cat.Name,
                    "Hidden" if set_hidden else "Visible"
                ])
        if errors:
            writer.writerow([])
            writer.writerow(["Warnings / Errors"])
            for msg in errors:
                writer.writerow([msg])
except Exception as e:
    forms.alert("Failed to write audit log:\n{}\nPath: {}".format(e, log_path))

forms.alert("Updates complete.\nAudit log saved to:\n{}".format(log_path))