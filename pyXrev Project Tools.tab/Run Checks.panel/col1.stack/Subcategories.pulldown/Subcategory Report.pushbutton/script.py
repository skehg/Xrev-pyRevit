# -*- coding: utf-8 -*-
# Create Family subcategory report (filtered to INVALID BuiltInCategory),
# build an in-memory list `project_only_subcats` of project-level INVALID subcategory
# Category objects that are NOT found in any family (normalized comparison),
# present a pyRevit UI to purge selected project-only subcategories, and
# open the Family CSV when finished.
#
# Only the Family CSV is written (for audit). The project CSV is not created.
#
# CSV encoding: Windows ANSI (mbcs)

import csv
import os
import codecs
import re
import subprocess
import unicodedata
import traceback

from Autodesk.Revit.DB import FilteredElementCollector, Family, CategoryType
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit
from pyrevit.forms import SelectFromList, pick_folder
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

def to_text(v):
    try:
        return unicode(v)   # noqa: F821 (IronPython)
    except NameError:
        return str(v)

def normalize_name_for_compare(value):
    """Robust normalization for comparison:
    - ensure text type
    - Unicode NFC
    - remove zero-width/BOM
    - replace underscores with spaces
    - collapse whitespace, strip, lower
    """
    if not value:
        return ""
    try:
        s = to_text(value)
    except Exception:
        try:
            s = str(value)
        except Exception:
            s = ""
    try:
        s = unicodedata.normalize("NFC", s)
    except Exception:
        pass
    for ch in ("\ufeff", "\u200b", "\u200c", "\u200d", "\uFEFF"):
        s = s.replace(ch, "")
    s = s.replace("_", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()

def codepoints(s):
    try:
        return " ".join("U+{:04X}".format(ord(ch)) for ch in to_text(s))
    except Exception:
        try:
            return " ".join("U+{:04X}".format(ord(ch)) for ch in str(s))
        except Exception:
            return ""

# Prompt where to save CSVs
csv_folder = pick_folder(title="Select folder to save the subcategories report")
if not csv_folder:
    script.exit()

# Get parent categories (model categories with no parent)
all_categories = doc.Settings.Categories
parent_categories = sorted(
    [cat.Name for cat in all_categories if cat.CategoryType == CategoryType.Model and cat.Parent is None]
)

# Prompt user to pick parent category
selected = SelectFromList.show(parent_categories, title="Select Parent Category", multiselect=False)
if not selected:
    TaskDialog.Show("Cancelled", "No category selected. Script aborted.")
    script.exit()

selected_cat_name = selected[0] if isinstance(selected, (list, tuple)) else selected
selected_cat = next((c for c in all_categories if c.Name == selected_cat_name), None)
if not selected_cat:
    TaskDialog.Show("Error", "Could not resolve selected category.")
    script.exit()

# CSV filename (Family only)
safe_name = re.sub(r'[<>:"/\\|?*\n\r\t]+', '_', selected_cat_name).strip().replace(' ', '_') or "Category"
family_csv_path = os.path.join(csv_folder, "{}_FamilySubcategoriesReport.csv".format(safe_name))

selected_cat_id = selected_cat.Id

# Part 1: scan families to collect family INVALID rows and normalized subcategory names
family_invalid_rows = []            # (FamilyName, SubcategoryName, BuiltInCategory)
family_subcat_names_norm = set()    # normalized subcategory names present in families

fam_collector = FilteredElementCollector(doc).OfClass(Family)
for fam in fam_collector:
    try:
        fam_cat = fam.FamilyCategory
    except Exception:
        fam_cat = None
    if not fam_cat or fam_cat.Id != selected_cat_id:
        continue

    fam_doc = None
    try:
        fam_doc = doc.EditFamily(fam)
        root_cat = fam_doc.OwnerFamily.FamilyCategory
        if not root_cat or root_cat.Id != selected_cat_id:
            continue

        for subcat in root_cat.SubCategories:
            # normalized name for comparison
            try:
                nm = subcat.Name
                if nm:
                    family_subcat_names_norm.add(normalize_name_for_compare(nm))
            except Exception:
                pass

            # BuiltInCategory: -1 => INVALID else enum string
            try:
                bic = subcat.BuiltInCategory
                bic_name = "INVALID" if bic == -1 else str(bic)
            except Exception:
                bic_name = "INVALID"

            if bic_name == "INVALID":
                family_invalid_rows.append((fam.Name, subcat.Name, bic_name))

    except Exception:
        # skip family if it fails to open
        continue
    finally:
        try:
            if fam_doc:
                fam_doc.Close(False)
        except Exception:
            pass

# Write only the family INVALID CSV
try:
    with codecs.open(family_csv_path, "w", encoding="mbcs") as f:
        writer = csv.writer(f)
        writer.writerow(["Family Name", "Subcategory Name", "BuiltInCategory", "New Subcategory"])
        for r in family_invalid_rows:
            writer.writerow(r)
    output.print_md("✅ Family INVALID CSV written: `{}`".format(family_csv_path))
except Exception as e:
    output.print_md("⚠️ Failed to write family CSV: {}".format(traceback.format_exc()))

# Part 2: list project subcategories where BuiltInCategory == INVALID and try to get GraphicStyleCategory
project_rows = []   # (SubcategoryName, BuiltInCategory, GraphicStyleCategory, NormalizedName, Codepoints, CategoryObj)
project_subcats = [sc for sc in selected_cat.SubCategories]  # snapshot

for sc in project_subcats:
    try:
        name = sc.Name if sc.Name else ""
    except Exception:
        name = ""
    try:
        bic = sc.BuiltInCategory
        bic_name = "INVALID" if bic == -1 else str(bic)
    except Exception:
        bic = None
        bic_name = "INVALID"

    # Only include INVALID entries
    if bic_name != "INVALID":
        continue

    gstyle_name = ""
    # try a few ways to extract graphic style/category information; APIs vary by Revit version
    try:
        g = getattr(sc, "GraphicsStyleCategory", None)
        if g:
            try:
                gstyle_name = g.Name
            except Exception:
                gstyle_name = str(g)
    except Exception:
        pass

    if not gstyle_name:
        try:
            gs = getattr(sc, "GraphicsStyle", None)
            if gs:
                try:
                    gstyle_name = gs.Name
                except Exception:
                    gstyle_name = str(gs)
        except Exception:
            pass

    if not gstyle_name:
        try:
            gid = getattr(sc, "GraphicsStyleId", None)
            if gid:
                try:
                    elem = doc.GetElement(gid)
                    if elem:
                        try:
                            gstyle_name = elem.Name
                        except Exception:
                            gstyle_name = str(elem)
                except Exception:
                    pass
        except Exception:
            pass

    norm = normalize_name_for_compare(name)
    cps = codepoints(name)
    project_rows.append((name, bic_name, gstyle_name, norm, cps, sc))

# NOTE: project CSV not written per request — project_rows is kept for comparison

# Part 3: compute project-only INVALID subcategories (not found in family normalized names)
project_only_subcats = []  # list of Category objects (project-only invalid subcats)
for name, bic_name, gstyle_name, norm, cps, sc in project_rows:
    if norm not in family_subcat_names_norm:
        project_only_subcats.append(sc)

# Summary to user
output.print_md(
    "✅ Report created for category **{}**:\n\n"
    "- Family INVALID CSV: `{}`\n\n"
    "Project-only INVALID count (in-memory): {}".format(
        selected_cat_name, family_csv_path, len(project_only_subcats)
    )
)

# If there are project-only subcategories, prompt user to select and purge using pyRevit UI
if project_only_subcats:
    class SubCategoryOption(forms.TemplateListItem):
        def __init__(self, subcategory):
            super(SubCategoryOption, self).__init__(subcategory)
        @property
        def name(self):
            try:
                return '{} --> {}'.format(self.item.Parent.Name, self.item.Name)
            except Exception:
                return '{}'.format(getattr(self.item, "Name", "<unnamed>"))

    if forms.alert('This tool will purge selected project-only subcategories from the model.\n\n'
                   'Proceed only if you have a backup. Continue?', yes=True, no=True):
        # Build selectable list
        options = [SubCategoryOption(x) for x in project_only_subcats]
        selected_options = forms.SelectFromList.show(options,
                                                    title='Select Project-only SubCategories to Purge',
                                                    button_name='Purge',
                                                    multiselect=True,
                                                    checked_only=True)
        if selected_options and forms.alert('Are you sure you want to permanently delete the selected subcategories?', yes=True, no=True):
            # Extract raw Category objects from selected TemplateListItems (safe extraction)
            elems_to_delete = []
            for opt in selected_options:
                try:
                    # TemplateListItem keeps original in .item
                    elem = getattr(opt, "item", None)
                    if elem is None:
                        # maybe user passed the raw Category object
                        elem = opt
                    if elem is not None:
                        elems_to_delete.append(elem)
                except Exception:
                    continue

            if elems_to_delete:
                try:
                    with revit.Transaction('Purge Project SubCategories'):
                        # Use pyRevit delete helper which reliably handles GraphicsStyle elements backing subcategories
                        revit.delete.delete_elements(elems_to_delete)
                    # Verify and report
                    deleted = []
                    failed = []
                    for sc in elems_to_delete:
                        try:
                            name = getattr(sc, "Name", "<unnamed>")
                            # if doc.GetElement returns None, element was removed
                            still = doc.GetElement(sc.Id)
                            if still is None:
                                deleted.append(name)
                            else:
                                failed.append((name, "Still present after delete attempt"))
                        except Exception as e:
                            failed.append((getattr(sc, "Name", "<unnamed>"), "Verification error: {}".format(e)))
                    if deleted:
                        for n in deleted:
                            output.print_md("🗑️ Deleted project subcategory: {}".format(n))
                    if failed:
                        output.print_md("⚠️ Failed to delete the following subcategories:")
                        for n, err in failed:
                            output.print_md("- {} : {}".format(n, err))
                except Exception:
                    output.print_md("⚠️ Purge transaction failed:\n{}".format(traceback.format_exc()))
            else:
                output.print_md("ℹ️ No elements selected for deletion.")
        else:
            output.print_md("Purge cancelled by user.")
else:
    output.print_md("✅ No project-only INVALID subcategories found to purge (no action).")

# Open the Family CSV (instead of opening the folder)
try:
    # Try os.startfile (Windows)
    os.startfile(family_csv_path)
except Exception:
    try:
        # Fallback: open with explorer select (if available)
        subprocess.Popen(['explorer', os.path.normpath(family_csv_path)])
    except Exception:
        try:
            script.open_url("file:///" + family_csv_path.replace("\\", "/"))
        except Exception:
            TaskDialog.Show("Notice", "Report created but could not be opened automatically.\nFile saved to:\n{}".format(family_csv_path))