# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, Transaction,
    CategoryType
)
from pyrevit import forms
import os

doc = __revit__.ActiveUIDocument.Document

# -------------------------------
# Collect view templates
# -------------------------------
templates = [v for v in FilteredElementCollector(doc).OfClass(View) if v.IsTemplate]
if not templates:
    forms.alert("No view templates found.", exitscript=True)

templates_sorted = sorted(templates, key=lambda v: v.Name)

# Select source template
source_name = forms.SelectFromList.show(
    [v.Name for v in templates_sorted],
    multiselect=False,
    title="Select Source View Template"
)
if not source_name:
    forms.alert("No source template selected.", exitscript=True)

source_vt = next(v for v in templates if v.Name == source_name)

# Select target templates
target_names = forms.SelectFromList.show(
    sorted([v.Name for v in templates if v.Name != source_name]),
    multiselect=True,
    title="Select Target View Templates"
)
if not target_names:
    forms.alert("No target templates selected.", exitscript=True)

target_vts = [v for v in templates if v.Name in target_names]

# Build overrideable parent category list
all_cats = doc.Settings.Categories
parents = [c for c in all_cats if c.Parent is None]

overrideable = []
for c in parents:
    allows_vis = False
    can_hide = False
    try:
        allows_vis = c.AllowsVisibilityControl(source_vt)
    except:
        pass
    try:
        can_hide = source_vt.CanCategoryBeHidden(c.Id)
    except:
        pass
    if allows_vis or can_hide:
        overrideable.append(c)

model_cats = sorted([c for c in overrideable if c.CategoryType == CategoryType.Model], key=lambda x: x.Name)
anno_cats  = sorted([c for c in overrideable if c.CategoryType == CategoryType.Annotation], key=lambda x: x.Name)

grouped_list = ["Model Categories"]
grouped_list.extend([c.Name for c in model_cats])
grouped_list.append("")
grouped_list.append("Annotation Categories")
grouped_list.extend([c.Name for c in anno_cats])

# -------------------------------
# Load WPF form from XAML file
# -------------------------------
xaml_path = os.path.join(os.path.dirname(__file__), "TransferForm.xaml")

class TransferForm(forms.WPFWindow):
    def __init__(self, xamlfile, categories):
        forms.WPFWindow.__init__(self, xamlfile)
        self.lstCategories.ItemsSource = categories
        self.btnOK.Click += self.on_ok
        self.btnCancel.Click += self.on_cancel
        self.selected = None
        self.copy_visibility = True
        self.copy_styles = True

    def on_ok(self, sender, args):
        self.selected = [item for item in self.lstCategories.SelectedItems if not item.startswith("===")]
        self.copy_visibility = self.chkVisibility.IsChecked
        self.copy_styles = self.chkStyles.IsChecked
        self.Close()

    def on_cancel(self, sender, args):
        self.Close()

form = TransferForm(xaml_path, grouped_list)
form.ShowDialog()

if not form.selected:
    forms.alert("No categories selected.", exitscript=True)

cats_to_transfer = [c for c in overrideable if c.Name in form.selected]

# -------------------------------
# Apply transfer
# -------------------------------
errors = []
t = Transaction(doc, "Transfer Category Settings")
try:
    t.Start()
    for tgt in target_vts:
        for cat in cats_to_transfer:
            try:
                if form.copy_styles:
                    ogs = source_vt.GetCategoryOverrides(cat.Id)
                    tgt.SetCategoryOverrides(cat.Id, ogs)
                if form.copy_visibility:
                    hidden = source_vt.GetCategoryHidden(cat.Id)
                    tgt.SetCategoryHidden(cat.Id, hidden)
            except Exception as e:
                errors.append("Error on '{}' in '{}': {}".format(cat.Name, tgt.Name, e))
    t.Commit()
except Exception as e:
    errors.append("Transaction failed: {}".format(e))
    try:
        t.RollBack()
    except:
        pass

msg = "Transferred settings from '{}' to {} templates.\nOptions: ".format(
    source_vt.Name, len(target_vts)
)
opts = []
if form.copy_styles: opts.append("Visual Styles")
if form.copy_visibility: opts.append("Visibility State")
msg += " + ".join(opts) if opts else "None"

if errors:
    msg += "\n\nErrors:\n" + "\n".join(errors)

forms.alert(msg)