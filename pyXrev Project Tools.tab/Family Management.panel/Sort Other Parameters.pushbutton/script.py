# -*- coding: utf-8 -*-
import clr
from pyrevit import revit, forms
import sys

clr.AddReference("PresentationFramework")
from System.Windows.Controls import CheckBox
from System.Windows import Thickness

from sort_param_utils import (
    get_current_parameter_order,
    supports_reorder,
    group_parameters,
    is_other_group,
    apply_parameter_sort,
)

doc = revit.doc


def _exit(message, title):
    forms.alert(message, title=title)
    sys.exit()


class SortSettingsWindow(forms.WPFWindow):
    def __init__(self, grouped, labels):
        forms.WPFWindow.__init__(self, "SortParameterGroupsForm.xaml")
        self.selected_keys = set()
        self.sort_mode = "Name Only"
        self.sort_direction = "Ascending (A-Z)"
        self._group_checkboxes = []

        for key in sorted(grouped.keys(), key=lambda k: labels.get(k, "").lower()):
            label = labels.get(key, key)
            cb = CheckBox()
            cb.Content = "{} ({})".format(label, len(grouped[key]))
            cb.Tag = key
            cb.Margin = Thickness(0, 2, 0, 2)
            cb.IsChecked = is_other_group(key, label)
            self.groupPanel.Children.Add(cb)
            self._group_checkboxes.append(cb)

    def on_select_all(self, sender, args):
        for cb in self._group_checkboxes:
            cb.IsChecked = True

    def on_select_none(self, sender, args):
        for cb in self._group_checkboxes:
            cb.IsChecked = False

    def on_ok(self, sender, args):
        selected = set()
        for cb in self._group_checkboxes:
            if bool(cb.IsChecked):
                selected.add(cb.Tag)

        if not selected:
            forms.alert("Select at least one parameter group.", title="No Groups Selected")
            return

        self.selected_keys = selected
        self.sort_mode = "Type then Instance" if bool(self.rbTypeThenInstance.IsChecked) else "Name Only"
        self.sort_direction = "Descending (Z-A)" if bool(self.rbDescending.IsChecked) else "Ascending (A-Z)"
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if not doc.IsFamilyDocument:
    _exit(
        "This tool only works in the Family Editor.\n"
        "Please open a family document (.rfa) and run again.",
        "Not a Family Document"
    )

fm = doc.FamilyManager
all_params = get_current_parameter_order(fm)

if all_params is None:
    _exit(
        "This Revit version does not expose FamilyManager.GetParameters, which is needed\n"
        "to preserve manual ordering in non-target groups.\n"
        "To avoid unintended reordering, this command has been cancelled.",
        "Unsupported API"
    )

if not all_params:
    _exit("No family parameters were found in this family.", "Nothing to Sort")

if not supports_reorder(fm):
    _exit(
        "This Revit version does not expose the FamilyManager.ReorderParameters API.\n"
        "The tool cannot apply an in-place group-only sort in this version.",
        "Unsupported API"
    )

grouped_params, group_labels = group_parameters(all_params)
dlg = SortSettingsWindow(grouped_params, group_labels)
if not dlg.ShowDialog():
    sys.exit()

selected_keys = dlg.selected_keys
sort_mode = dlg.sort_mode
descending = (dlg.sort_direction == "Descending (Z-A)")

success, result_code, changed_group_labels, processed_count = apply_parameter_sort(
    doc, fm, selected_keys, sort_mode, descending
)

if not success:
    _exit(
        "Failed to reorder selected parameter groups:\n\n{}".format(result_code),
        "Sort Failed"
    )

if result_code == "already_sorted":
    forms.alert("Selected groups are already in alphabetical order.", title="No Changes Needed")
    sys.exit()

if processed_count < 2:
    forms.alert(
        "Found {} parameter(s) across selected group(s).\n"
        "At least 2 are needed to change order.".format(processed_count),
        title="Nothing to Reorder"
    )
    sys.exit()

forms.alert(
    "Sorted selected groups alphabetically.\n"
    "Sort mode: {}\n"
    "Sort direction: {}\n"
    "Groups changed: {}\n"
    "Parameters considered in selected groups: {}\n"
    "All non-selected groups were kept in their original order.".format(
        sort_mode,
        dlg.sort_direction,
        ", ".join(sorted(changed_group_labels)),
        processed_count
    ),
    title="Sort Complete"
)
