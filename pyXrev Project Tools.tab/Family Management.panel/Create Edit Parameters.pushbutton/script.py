# -*- coding: utf-8 -*-
"""
Parameter Editor
----------------
Create family/shared parameters and edit formulas for existing parameters.
Formula autocomplete suggests existing family parameter names while typing.
"""

import re
import sys
import clr
import System

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (
    BuiltInParameterGroup,
    ElementId,
    FilteredElementCollector,
    LabelUtils,
    Material,
    ParameterType,
    StorageType,
    Transaction,
    UnitUtils,
)
from pyrevit import forms, revit
from family_param_utils import (
    find_directly_used_params,
    find_formula_referencing_params,
)
from sort_param_utils import (
    get_current_parameter_order as _sort_get_current_order,
    group_parameters as _sort_group_parameters,
    is_other_group as _sort_is_other_group,
    supports_reorder as _sort_supports_reorder,
    apply_parameter_sort as _sort_apply_parameter_sort,
)
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import CheckBox
from System.Windows import Thickness
from System.Windows.Input import Key


doc = revit.doc
uiapp = __revit__  # noqa: F821
app = uiapp.Application


class OptionItem(object):
    def __init__(self, label, value):
        self.Label = label
        self.Value = value

    def __str__(self):
        return self.Label


class ParameterItem(object):
    def __init__(self, family_param, current_type):
        self.Param = family_param
        self.Name = _param_name(family_param)
        self.GroupLabel = _group_label(family_param)
        self.InstanceTypeLabel = "Instance" if bool(getattr(family_param, "IsInstance", False)) else "Type"
        self.DataTypeLabel = _parameter_data_type_label(family_param)
        self.Formula = _safe_formula(family_param)
        self.IsShared = _is_shared_parameter(family_param)
        self.Value = _read_value_for_display(current_type, family_param)


class ParameterEditorWindow(forms.WPFWindow):
    def __init__(self, family_manager):
        forms.WPFWindow.__init__(self, "ParameterEditor.xaml")
        self.fm = family_manager
        self._all_groups_label = "All Groups"
        self._all_items = []
        self._filtered_items = ObservableCollection[object]()
        self._autocomplete_names = []
        self._shared_options = []
        self._active_token_range = None

        self.lstParameters.ItemsSource = self._filtered_items

        self._type_options = _build_type_options(self.fm)
        self._group_options = _build_group_options(self.fm)
        self._shared_options = _build_shared_definition_options()

        self._bind_combo_options(self.cmbNewType, self._type_options)
        self._bind_combo_options(self.cmbNewGroup, self._group_options)
        self._bind_combo_options(self.cmbSharedDefinition, self._shared_options)
        self._bind_combo_options(self.cmbEditGroup, self._group_options)

        self._set_shared_mode(False)
        self._reload_parameter_items(select_name=None)
        self._sort_group_checkboxes = []
        self._populate_sort_groups()
        self._restore_window_position()
        self._set_status("Ready.", "neutral")

    def _bind_combo_options(self, combo, options):
        combo.ItemsSource = options
        combo.DisplayMemberPath = "Label"
        combo.SelectedValuePath = "Value"
        if options:
            combo.SelectedIndex = 0

    def _set_status(self, message, tone):
        self.txtStatus.Text = message or ""
        if tone == "error":
            self.txtStatus.Foreground = self._brush("#B00020")
        elif tone == "ok":
            self.txtStatus.Foreground = self._brush("#1B5E20")
        else:
            self.txtStatus.Foreground = self._brush("#2A2A2A")

    def _brush(self, hex_color):
        from System.Windows.Media import BrushConverter

        return BrushConverter().ConvertFromString(hex_color)

    def _uniform_thickness(self, size):
        from System.Windows import Thickness

        return Thickness(size)

    def _reload_parameter_items(self, select_name=None):
        current_type = self.fm.CurrentType
        self._all_items = [ParameterItem(p, current_type) for p in _get_family_parameters(self.fm)]
        self._autocomplete_names = sorted({item.Name for item in self._all_items if item.Name})
        self._refresh_group_filter_options()
        self._apply_parameter_filter()

        if not self._all_items:
            self.lstParameters.SelectedItem = None
            self._set_editor_from_selected()
            return

        selected = None
        if select_name:
            for item in self._all_items:
                if item.Name == select_name:
                    selected = item
                    break
        if selected is None:
            selected = self._all_items[0]
        self.lstParameters.SelectedItem = selected

    def _apply_parameter_filter(self):
        search = (self.txtSearch.Text or "").strip().lower()
        selected_group = self.cmbGroupFilter.SelectedItem
        group_filter = None
        if selected_group and selected_group != self._all_groups_label:
            group_filter = str(selected_group).strip().lower()

        self._filtered_items.Clear()
        for item in self._all_items:
            hay_name = (item.Name or "").lower()
            hay_formula = (item.Formula or "").lower()
            hay_group = (item.GroupLabel or "").strip().lower()

            if group_filter and hay_group != group_filter:
                continue
            if search and search not in hay_name and search not in hay_formula:
                continue

            self._filtered_items.Add(item)

    def _refresh_group_filter_options(self):
        current_selection = self.cmbGroupFilter.SelectedItem
        groups = sorted({
            (item.GroupLabel or "").strip()
            for item in self._all_items
            if (item.GroupLabel or "").strip()
        })
        options = [self._all_groups_label] + groups
        self.cmbGroupFilter.ItemsSource = options

        if current_selection in options:
            self.cmbGroupFilter.SelectedItem = current_selection
        elif options:
            self.cmbGroupFilter.SelectedIndex = 0

    def _set_editor_from_selected(self):
        item = self.lstParameters.SelectedItem
        if item is None:
            self.txtSelectedParamName.Text = ""
            self.txtSelectedParamDataType.Text = ""
            self.txtSelectedParamValue.Text = ""
            self.txtRenameTo.Text = ""
            self.txtFormula.Text = ""
            self.txtFormulaBracketStatus.Text = ""
            self.txtInstanceType.Text = ""
            self.txtFormula.BorderBrush = self._brush("#B5B5B5")
            self.txtFormula.BorderThickness = self._uniform_thickness(1)
            return

        self.txtSelectedParamName.Text = item.Name
        self.txtSelectedParamDataType.Text = item.DataTypeLabel
        self.txtSelectedParamValue.Text = item.Value
        self.txtRenameTo.Text = item.Name
        self.txtFormula.Text = item.Formula or ""
        self._set_edit_group_selection(item.Param)
        self.txtInstanceType.Text = item.InstanceTypeLabel
        self._update_formula_bracket_feedback()

    def _set_edit_group_selection(self, family_param):
        try:
            current_group = family_param.Definition.ParameterGroup
        except Exception:
            current_group = None

        if current_group is None:
            return

        for opt in self._group_options:
            if str(opt.Value) == str(current_group):
                self.cmbEditGroup.SelectedItem = opt
                return

    def _set_shared_mode(self, is_shared):
        self.chkNewShared.IsChecked = is_shared
        self.cmbSharedDefinition.IsEnabled = is_shared
        self.cmbNewType.IsEnabled = not is_shared
        self.txtNewName.IsEnabled = not is_shared

        if is_shared and self._shared_options and self.cmbSharedDefinition.SelectedItem is None:
            self.cmbSharedDefinition.SelectedIndex = 0

    def _selected_token_range(self, text, caret):
        if text is None:
            return None

        n = len(text)
        if caret < 0:
            caret = 0
        if caret > n:
            caret = n

        left = caret - 1
        while left >= 0 and _is_token_char(text[left]):
            left -= 1
        left += 1

        right = caret
        while right < n and _is_token_char(text[right]):
            right += 1

        if left >= right:
            return None
        return (left, right)

    def _render_suggestions(self):
        txt = self.txtFormula.Text or ""
        rng = self._selected_token_range(txt, self.txtFormula.CaretIndex)
        self._active_token_range = rng

        if rng is None:
            self.popupSuggestions.IsOpen = False
            return

        token = txt[rng[0]:rng[1]].strip()
        if not token:
            self.popupSuggestions.IsOpen = False
            return

        token_lower = token.lower()
        selected = self.lstParameters.SelectedItem
        selected_name = selected.Name if selected else None

        prefix = []
        contains = []
        for name in self._autocomplete_names:
            if not name:
                continue
            if selected_name and name == selected_name:
                continue
            low = name.lower()
            if low.startswith(token_lower):
                prefix.append(name)
            elif token_lower in low:
                contains.append(name)

        suggestions = (prefix + contains)[:40]
        self.lstSuggestions.ItemsSource = suggestions
        if suggestions:
            self.lstSuggestions.SelectedIndex = 0
            self.popupSuggestions.IsOpen = True
        else:
            self.popupSuggestions.IsOpen = False

    def _commit_selected_suggestion(self):
        if not self.popupSuggestions.IsOpen:
            return False

        suggestion = self.lstSuggestions.SelectedItem
        if suggestion is None:
            return False

        text = self.txtFormula.Text or ""
        rng = self._active_token_range or self._selected_token_range(text, self.txtFormula.CaretIndex)
        if rng is None:
            return False

        new_text = text[:rng[0]] + suggestion + text[rng[1]:]
        caret_pos = rng[0] + len(suggestion)
        self.txtFormula.Text = new_text
        self.popupSuggestions.IsOpen = False

        # Keep typing flow smooth: focus formula box and place caret after inserted suggestion.
        self.txtFormula.Focus()
        self.txtFormula.CaretIndex = caret_pos
        self.txtFormula.SelectionStart = caret_pos
        self.txtFormula.SelectionLength = 0
        return True

    def _insert_pair(self, opener, closer):
        tb = self.txtFormula
        text = tb.Text or ""
        start = tb.SelectionStart
        length = tb.SelectionLength

        if length > 0:
            selected = text[start:start + length]
            new_text = text[:start] + opener + selected + closer + text[start + length:]
            tb.Text = new_text
            tb.SelectionStart = start + length + 2
            tb.SelectionLength = 0
            return

        new_text = text[:start] + opener + closer + text[start:]
        tb.Text = new_text
        tb.SelectionStart = start + 1
        tb.SelectionLength = 0

    def _find_matching_bracket(self, text, index):
        if index < 0 or index >= len(text):
            return None

        ch = text[index]
        open_to_close = {"(": ")", "[": "]", "{": "}"}
        close_to_open = {")": "(", "]": "[", "}": "{"}

        if ch in open_to_close:
            target = open_to_close[ch]
            depth = 0
            i = index + 1
            while i < len(text):
                c = text[i]
                if c == ch:
                    depth += 1
                elif c == target:
                    if depth == 0:
                        return i
                    depth -= 1
                i += 1
            return None

        if ch in close_to_open:
            target = close_to_open[ch]
            depth = 0
            i = index - 1
            while i >= 0:
                c = text[i]
                if c == ch:
                    depth += 1
                elif c == target:
                    if depth == 0:
                        return i
                    depth -= 1
                i -= 1
            return None

        return None

    def _find_first_unmatched_bracket(self, text):
        open_to_close = {"(": ")", "[": "]", "{": "}"}
        close_to_open = {")": "(", "]": "[", "}": "{"}
        stack = []

        for idx, ch in enumerate(text):
            if ch in open_to_close:
                stack.append((ch, idx))
                continue

            if ch in close_to_open:
                if not stack:
                    return ("unmatched_closer", ch, idx)
                top_ch, top_idx = stack[-1]
                if top_ch == close_to_open[ch]:
                    stack.pop()
                else:
                    return ("mismatched_closer", ch, idx)

        if stack:
            top_ch, top_idx = stack[-1]
            return ("unmatched_opener", top_ch, top_idx)

        return None

    def _bracket_color_hex(self, ch):
        if ch in ["(", ")"]:
            return "#1565C0"
        if ch in ["[", "]"]:
            return "#2E7D32"
        if ch in ["{", "}"]:
            return "#EF6C00"
        return "#6A6A6A"

    def _set_formula_feedback(self, message, color_hex, border_hex=None, border_size=1):
        self.txtFormulaBracketStatus.Text = message or ""
        self.txtFormulaBracketStatus.Foreground = self._brush(color_hex)

        if border_hex is None:
            border_hex = "#B5B5B5"
        self.txtFormula.BorderBrush = self._brush(border_hex)
        self.txtFormula.BorderThickness = self._uniform_thickness(border_size)

    def _update_formula_bracket_feedback(self):
        text = self.txtFormula.Text or ""
        caret = self.txtFormula.CaretIndex
        bracket_chars = set(["(", ")", "[", "]", "{", "}"])

        near_idx = None
        if caret > 0 and text[caret - 1] in bracket_chars:
            near_idx = caret - 1
        elif caret < len(text) and text[caret] in bracket_chars:
            near_idx = caret

        if near_idx is not None:
            ch = text[near_idx]
            match_idx = self._find_matching_bracket(text, near_idx)
            if match_idx is not None:
                color = self._bracket_color_hex(ch)
                self._set_formula_feedback(
                    "Bracket pair '{}' matched (positions {} and {}).".format(ch, near_idx + 1, match_idx + 1),
                    color,
                    color,
                    2,
                )
                return

            self._set_formula_feedback(
                "Bracket '{}' at position {} has no matching pair.".format(ch, near_idx + 1),
                "#B00020",
                "#B00020",
                2,
            )
            return

        unmatched = self._find_first_unmatched_bracket(text)
        if unmatched is not None:
            kind, ch, idx = unmatched
            if kind == "mismatched_closer":
                msg = "Mismatched closing bracket '{}' at position {}.".format(ch, idx + 1)
            elif kind == "unmatched_closer":
                msg = "Unmatched closing bracket '{}' at position {}.".format(ch, idx + 1)
            else:
                msg = "Unmatched opening bracket '{}' at position {}.".format(ch, idx + 1)

            self._set_formula_feedback(msg, "#B00020", "#B00020", 2)
            return

        if not text:
            self._set_formula_feedback("", "#6A6A6A", "#B5B5B5", 1)
            return

        self._set_formula_feedback("Brackets are balanced.", "#6A6A6A", "#B5B5B5", 1)

    def _selected_group_value(self):
        item = self.cmbNewGroup.SelectedItem
        return item.Value if item else None

    def _selected_type_value(self):
        item = self.cmbNewType.SelectedItem
        return item.Value if item else None

    def _selected_shared_definition(self):
        item = self.cmbSharedDefinition.SelectedItem
        return item.Value if item else None

    def _selected_parameter_item(self):
        return self.lstParameters.SelectedItem

    def _selected_edit_group_value(self):
        item = self.cmbEditGroup.SelectedItem
        return item.Value if item else None

    def on_search_changed(self, sender, args):
        self._apply_parameter_filter()

    def on_group_filter_changed(self, sender, args):
        self._apply_parameter_filter()

    def on_param_selected(self, sender, args):
        self._set_editor_from_selected()
        self.popupSuggestions.IsOpen = False

    def on_formula_changed(self, sender, args):
        self._render_suggestions()
        self._update_formula_bracket_feedback()

    def on_formula_selection_changed(self, sender, args):
        self._update_formula_bracket_feedback()

    def on_formula_preview_textinput(self, sender, args):
        # Bracket auto-close is intentionally disabled.
        return

    def on_formula_keydown(self, sender, args):
        if args.Key == Key.Down and self.popupSuggestions.IsOpen:
            idx = self.lstSuggestions.SelectedIndex
            if idx < self.lstSuggestions.Items.Count - 1:
                self.lstSuggestions.SelectedIndex = idx + 1
                self.lstSuggestions.ScrollIntoView(self.lstSuggestions.SelectedItem)
            args.Handled = True
            return

        if args.Key == Key.Up and self.popupSuggestions.IsOpen:
            idx = self.lstSuggestions.SelectedIndex
            if idx > 0:
                self.lstSuggestions.SelectedIndex = idx - 1
                self.lstSuggestions.ScrollIntoView(self.lstSuggestions.SelectedItem)
            args.Handled = True
            return

        if args.Key == Key.Enter and self.popupSuggestions.IsOpen:
            if self._commit_selected_suggestion():
                args.Handled = True
            return

        if args.Key == Key.Tab and self.popupSuggestions.IsOpen:
            if self._commit_selected_suggestion():
                args.Handled = True
            return

        if args.Key == Key.Escape and self.popupSuggestions.IsOpen:
            self.popupSuggestions.IsOpen = False
            args.Handled = True

    def on_suggestion_double_click(self, sender, args):
        self._commit_selected_suggestion()

    def on_apply_param_settings(self, sender, args):
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        target_group = self._selected_edit_group_value()
        if target_group is None:
            self._set_status("Select a group.", "error")
            return

        source_param = item.Param
        source_group = source_param.Definition.ParameterGroup

        if str(source_group) == str(target_group):
            self._set_status("Parameter is already in that group.", "neutral")
            return

        move_method = getattr(self.fm, "MoveParameter", None)
        if not callable(move_method):
            self._set_status("Group move requires Revit 2023 or later.", "error")
            return

        try:
            with revit.Transaction("Move Parameter Group"):
                move_method(source_param, target_group)
        except Exception as ex:
            self._set_status("Move group failed: {}".format(ex), "error")
            return

        self._set_status("Moved '{}' to group '{}'.".format(item.Name, _label_for_group(target_group)), "ok")
        self._reload_parameter_items(select_name=item.Name)

    def on_duplicate_parameter(self, sender, args):
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        source_param = item.Param
        new_name = _next_duplicate_name(item.Name, [it.Name for it in self._all_items])
        created_as_shared = False

        if item.IsShared:
            choice = forms.alert(
                "Shared parameters cannot be duplicated with a new name using the same shared definition.\n"
                "This duplicate will be created as a non-shared family parameter instead.\n\n"
                "Do you want to continue?",
                title="Duplicate Shared Parameter",
                options=["Continue", "Cancel"],
            )
            if choice != "Continue":
                self._set_status("Duplicate cancelled.", "neutral")
                return

        try:
            with revit.Transaction("Duplicate Family Parameter"):
                new_param = _duplicate_family_parameter(self.fm, source_param, new_name)
                created_as_shared = _is_shared_parameter(new_param)

                if not _safe_formula(source_param):
                    _copy_current_parameter_value(self.fm, source_param, new_param)

                source_formula = _safe_formula(source_param)
                if source_formula:
                    self.fm.SetFormula(new_param, source_formula)
        except Exception as ex:
            self._set_status("Duplicate failed: {}".format(ex), "error")
            return

        if item.IsShared and not created_as_shared:
            self._set_status(
                "Duplicated '{}' as non-shared parameter '{}'.".format(item.Name, new_name),
                "ok",
            )
        else:
            self._set_status("Duplicated parameter '{}' as '{}'.".format(item.Name, new_name), "ok")

        self._reload_parameter_items(select_name=new_name)
        self._populate_sort_groups()

    def on_delete_parameter(self, sender, args):
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        name = item.Name

        # Warn if parameter is actively used (dimensions, associations, arrays)
        try:
            used_set = find_directly_used_params(doc, self.fm)
            is_used = any(_param_name(fp) == name for fp in used_set)
        except Exception:
            is_used = False

        # Warn if other parameters reference this one in their formulas
        referencing = find_formula_referencing_params(self.fm, name)

        # Build warning message
        warning_lines = [
            "Delete parameter '{}'?\n\nTo undo, close the Parameter Editor and use Revit's Undo (Ctrl+Z).".format(name)
        ]
        if is_used:
            warning_lines.append(
                "\u26a0  This parameter appears to be IN USE (dimension label, "
                "element association, or array count).  Deleting it may break "
                "the family geometry."
            )
        if referencing:
            warning_lines.append(
                "\u26a0  The following parameter(s) reference '{}' in their "
                "formulas and will be broken:\n    {}".format(
                    name, ", ".join(referencing)
                )
            )

        confirmed = forms.alert(
            "\n\n".join(warning_lines),
            title="Delete Parameter",
            ok=True,
            cancel=True,
        )
        if not confirmed:
            self._set_status("Delete cancelled.", "neutral")
            return

        try:
            with revit.Transaction("Delete Family Parameter"):
                self.fm.RemoveParameter(item.Param)
        except Exception as ex:
            self._set_status("Delete failed: {}".format(ex), "error")
            return

        self._set_status("Deleted parameter '{}'.".format(name), "ok")
        self._reload_parameter_items(select_name=None)
        self._populate_sort_groups()

    def on_apply_value(self, sender, args):
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        raw_value = self.txtSelectedParamValue.Text
        raw_value = raw_value if raw_value is not None else ""

        try:
            with revit.Transaction("Set Parameter Value"):
                ok, reason = _apply_initial_value(self.fm, item.Param, raw_value)
                if not ok:
                    raise Exception(reason)
        except Exception as ex:
            self._set_status("Value update failed: {}".format(ex), "error")
            return

        self._set_status("Value updated for '{}'.".format(item.Name), "ok")
        self._reload_parameter_items(select_name=item.Name)

    def on_apply_formula(self, sender, args):
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        formula_text = (self.txtFormula.Text or "").strip()
        formula_value = formula_text if formula_text else None

        txn = None
        try:
            txn = Transaction(doc, "Set Parameter Formula")
            txn.Start()
            self.fm.SetFormula(item.Param, formula_value)
            _force_family_recalc(doc)
            _force_family_type_cycle(self.fm)
            _force_family_recalc(doc)
            txn.Commit()
        except Exception as ex:
            try:
                if txn is not None and txn.GetStatus().ToString() == "Started":
                    txn.RollBack()
            except Exception:
                pass
            self._set_status("Formula update failed: {}".format(ex), "error")
            return

        self._set_status("Formula updated for '{}'.".format(item.Name), "ok")
        self._reload_parameter_items(select_name=item.Name)

    def on_clear_formula(self, sender, args):
        self.txtFormula.Text = ""

    def on_rename_parameter(self, sender, args):
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        new_name = (self.txtRenameTo.Text or "").strip()
        if not new_name:
            self._set_status("Enter a new parameter name.", "error")
            return

        old_name = item.Name
        if new_name == old_name:
            self._set_status("New name matches the current name. No change made.", "neutral")
            return

        existing_names = {it.Name.lower() for it in self._all_items if it.Name and it.Name != old_name}
        if new_name.lower() in existing_names:
            self._set_status("A parameter named '{}' already exists.".format(new_name), "error")
            return

        if item.IsShared:
            choice = forms.alert(
                "You are renaming a shared parameter.\n"
                "This can break schedules/tags or mappings that rely on the old name.\n\n"
                "Do you want to continue?",
                title="Rename Shared Parameter",
                options=["Rename", "Cancel"],
            )
            if choice != "Rename":
                self._set_status("Rename cancelled.", "neutral")
                return

        rename_method = getattr(self.fm, "RenameParameter", None)
        if not callable(rename_method):
            self._set_status("This Revit version does not expose FamilyManager.RenameParameter.", "error")
            return

        try:
            with revit.Transaction("Rename Family Parameter"):
                rename_method(item.Param, new_name)
        except Exception as ex:
            self._set_status("Rename failed: {}".format(ex), "error")
            return

        self._set_status("Renamed parameter '{}' to '{}'.".format(old_name, new_name), "ok")
        self._reload_parameter_items(select_name=new_name)

    def on_new_shared_changed(self, sender, args):
        self._set_shared_mode(bool(self.chkNewShared.IsChecked))

    def on_shared_def_selected(self, sender, args):
        if not bool(self.chkNewShared.IsChecked):
            return
        definition = self._selected_shared_definition()
        if definition is not None:
            self.txtNewName.Text = getattr(definition, "Name", "")

    def on_create_parameter(self, sender, args):
        is_shared = bool(self.chkNewShared.IsChecked)
        is_instance = bool(self.chkNewInstance.IsChecked)

        existing_names = {item.Name.lower() for item in self._all_items if item.Name}

        if is_shared:
            shared_def = self._selected_shared_definition()
            if shared_def is None:
                self._set_status("Select a shared parameter definition.", "error")
                return
            new_name = getattr(shared_def, "Name", "")
        else:
            shared_def = None
            new_name = (self.txtNewName.Text or "").strip()

        if not new_name:
            self._set_status("Enter a parameter name.", "error")
            return

        if new_name.lower() in existing_names:
            self._set_status("A parameter named '{}' already exists.".format(new_name), "error")
            return

        group_value = self._selected_group_value()
        if group_value is None:
            self._set_status("Select a parameter group.", "error")
            return

        type_value = self._selected_type_value()
        if (not is_shared) and type_value is None:
            self._set_status("Select a parameter type.", "error")
            return

        initial_value_text = (self.txtNewInitialValue.Text or "").strip()
        initial_formula_text = (self.txtNewFormula.Text or "").strip()

        try:
            with revit.Transaction("Create Family Parameter"):
                if is_shared:
                    new_param = self.fm.AddParameter(shared_def, group_value, is_instance)
                else:
                    new_param = self.fm.AddParameter(new_name, group_value, type_value, is_instance)

                if initial_value_text:
                    ok, reason = _apply_initial_value(self.fm, new_param, initial_value_text)
                    if not ok:
                        raise Exception(reason)

                if initial_formula_text:
                    self.fm.SetFormula(new_param, initial_formula_text)
        except Exception as ex:
            self._set_status("Create parameter failed: {}".format(ex), "error")
            return

        self.txtNewInitialValue.Text = ""
        self.txtNewFormula.Text = ""
        if not is_shared:
            self.txtNewName.Text = ""

        self._set_status("Created parameter '{}'.".format(new_name), "ok")
        self._reload_parameter_items(select_name=new_name)
        self._populate_sort_groups()

    # ------------------------------------------------------------------
    # Sort Parameters tab
    # ------------------------------------------------------------------

    def _populate_sort_groups(self):
        """Rebuild the group checkbox list in the Sort Parameters tab."""
        self.sortGroupPanel.Children.Clear()
        self._sort_group_checkboxes = []

        if not _sort_supports_reorder(self.fm):
            from System.Windows.Controls import TextBlock as _WpfTextBlock
            tb = _WpfTextBlock()
            tb.Text = "ReorderParameters API is not available in this Revit version."
            tb.TextWrapping = System.Windows.TextWrapping.Wrap
            self.sortGroupPanel.Children.Add(tb)
            return

        all_params = _sort_get_current_order(self.fm)
        if not all_params:
            return

        grouped, labels = _sort_group_parameters(all_params)
        for key in sorted(grouped.keys(), key=lambda k: labels.get(k, "").lower()):
            label = labels.get(key, key)
            cb = CheckBox()
            cb.Content = "{} ({})".format(label, len(grouped[key]))
            cb.Tag = key
            cb.Margin = Thickness(0, 2, 0, 2)
            cb.IsChecked = _sort_is_other_group(key, label)
            self.sortGroupPanel.Children.Add(cb)
            self._sort_group_checkboxes.append(cb)

    def on_sort_refresh(self, sender, args):
        self._populate_sort_groups()
        self._set_status("Group list refreshed.", "neutral")

    def on_sort_select_all(self, sender, args):
        for cb in self._sort_group_checkboxes:
            cb.IsChecked = True

    def on_sort_select_none(self, sender, args):
        for cb in self._sort_group_checkboxes:
            cb.IsChecked = False

    def on_sort_apply(self, sender, args):
        selected_keys = set(
            cb.Tag for cb in self._sort_group_checkboxes if bool(cb.IsChecked)
        )
        if not selected_keys:
            self._set_status("Select at least one group to sort.", "error")
            return

        sort_mode = "Type then Instance" if bool(self.rbSortTypeThenInstance.IsChecked) else "Name Only"
        descending = bool(self.rbSortDescending.IsChecked)

        success, result_code, changed_labels, count = _sort_apply_parameter_sort(
            doc, self.fm, selected_keys, sort_mode, descending
        )

        if not success:
            self._set_status("Sort failed: {}".format(result_code), "error")
            return

        if result_code == "already_sorted":
            self._set_status("Selected groups are already in alphabetical order.", "neutral")
            return

        self._set_status(
            "Sorted {} group(s): {}.".format(len(changed_labels), ", ".join(changed_labels)),
            "ok",
        )
        selected_name = None
        item = self._selected_parameter_item()
        if item:
            selected_name = item.Name
        self._reload_parameter_items(select_name=selected_name)
        self._populate_sort_groups()

    def _restore_window_position(self):
        from pyrevit import script as _pyscript
        from System.Windows import SystemParameters
        cfg = _pyscript.get_config()
        try:
            width = getattr(cfg, 'win_width', None)
            height = getattr(cfg, 'win_height', None)
            if width is not None:
                self.Width = float(width)
                self.Height = float(height)
        except Exception:
            pass
        try:
            left = getattr(cfg, 'win_left', None)
            top = getattr(cfg, 'win_top', None)
            if left is not None:
                left = float(left)
                top = float(top)
                vl = SystemParameters.VirtualScreenLeft
                vt = SystemParameters.VirtualScreenTop
                vr = vl + SystemParameters.VirtualScreenWidth
                vb = vt + SystemParameters.VirtualScreenHeight
                margin = 50.0
                if (left + margin < vr and left + self.Width - margin > vl
                        and top + margin < vb and top + margin > vt):
                    from System.Windows import WindowStartupLocation
                    self.WindowStartupLocation = WindowStartupLocation.Manual
                    self.Left = left
                    self.Top = top
        except Exception:
            pass

    def on_close(self, sender, args):
        from pyrevit import script as _pyscript
        cfg = _pyscript.get_config()
        cfg.win_left = self.Left
        cfg.win_top = self.Top
        cfg.win_width = self.Width
        cfg.win_height = self.Height
        _pyscript.save_config()
        self.Close()


def _is_token_char(ch):
    return ch.isalnum() or ch == "_"


def _get_family_parameters(fm):
    getter = getattr(fm, "GetParameters", None)
    if callable(getter):
        try:
            return list(getter())
        except Exception:
            pass

    try:
        return list(fm.Parameters)
    except Exception:
        return []


def _safe_formula(fp):
    try:
        return fp.Formula or ""
    except Exception:
        return ""


def _param_name(fp):
    try:
        return fp.Definition.Name or ""
    except Exception:
        return ""


def _is_shared_parameter(fp):
    try:
        return bool(fp.IsShared)
    except Exception:
        return False


def _read_value_for_display(current_type, fp):
    """Return a human-readable value string for the parameter in the current type."""
    if current_type is None:
        return "<No current type>"

    storage = getattr(fp, "StorageType", None)

    # Prefer Revit's formatted display string when available.
    try:
        val = current_type.AsValueString(fp)
        if val is not None:
            trimmed = val.strip() if hasattr(val, "strip") else val
            if trimmed:
                return trimmed
    except Exception:
        pass

    try:
        if storage == StorageType.String:
            s = current_type.AsString(fp)
            return s if s is not None else ""

        if storage == StorageType.Integer:
            i = current_type.AsInteger(fp)
            if _is_yesno_parameter(fp):
                return "Yes" if int(i) != 0 else "No"
            return str(i)

        if storage == StorageType.Double:
            d = current_type.AsDouble(fp)
            return str(d)

        if storage == StorageType.ElementId:
            eid = current_type.AsElementId(fp)
            if eid is None or eid == ElementId.InvalidElementId:
                return "<None>"

            elem = doc.GetElement(eid)
            if elem is not None:
                try:
                    name = getattr(elem, "Name", None)
                    if name:
                        return name
                except Exception:
                    pass

            try:
                return "ElementId({})".format(eid.IntegerValue)
            except Exception:
                return str(eid)
    except Exception:
        pass

    return ""


def _group_label(fp):
    try:
        g = fp.Definition.ParameterGroup
    except Exception:
        return ""

    try:
        return LabelUtils.GetLabelFor(g)
    except Exception:
        try:
            return str(g)
        except Exception:
            return ""


def _pretty_type_text(text):
    if not text:
        return "Unknown"

    mapping = {
        "YesNo": "Yes/No",
        "yesno": "Yes/No",
    }
    if text in mapping:
        return mapping[text]
    return text


def _parameter_data_type_label(fp):
    definition = getattr(fp, "Definition", None)
    if definition is None:
        return "Unknown"

    get_data_type = getattr(definition, "GetDataType", None)
    if callable(get_data_type):
        try:
            spec = get_data_type()
            if spec is not None:
                get_label_for_spec = getattr(LabelUtils, "GetLabelForSpec", None)
                if callable(get_label_for_spec):
                    try:
                        label = get_label_for_spec(spec)
                        if label:
                            return label
                    except Exception:
                        pass
                return _pretty_type_text(str(spec))
        except Exception:
            pass

    try:
        return _pretty_type_text(str(definition.ParameterType))
    except Exception:
        return "Unknown"


def _get_data_type(definition):
    try:
        return definition.GetDataType()
    except Exception:
        return getattr(definition, "ParameterType", None)


def _build_type_options(fm):
    options = []
    seen = set()

    for fp in _get_family_parameters(fm):
        try:
            dtype = _get_data_type(fp.Definition)
        except Exception:
            continue

        if dtype is None:
            continue

        key = str(dtype)
        if key in seen:
            continue
        seen.add(key)
        options.append(OptionItem(_parameter_data_type_label(fp), dtype))

    if options:
        options.sort(key=lambda o: o.Label.lower())
        return options

    try:
        fallback = [
            ParameterType.Text,
            ParameterType.Number,
            ParameterType.Length,
            ParameterType.YesNo,
            ParameterType.Integer,
        ]
        for item in fallback:
            options.append(OptionItem(_pretty_type_text(str(item)), item))
    except Exception:
        pass

    return options


def _build_group_options(fm):
    options = []
    seen = set()

    for fp in _get_family_parameters(fm):
        try:
            group = fp.Definition.ParameterGroup
        except Exception:
            continue

        key = str(group)
        if key in seen:
            continue
        seen.add(key)

        label = _label_for_group(group)
        options.append(OptionItem(label, group))

    if options:
        options.sort(key=lambda o: o.Label.lower())
        return options

    try:
        for group in System.Enum.GetValues(BuiltInParameterGroup):
            if str(group) == "INVALID":
                continue
            options.append(OptionItem(_label_for_group(group), group))
    except Exception:
        pass

    return options


def _label_for_group(group):
    try:
        return LabelUtils.GetLabelFor(group)
    except Exception:
        return str(group)


def _build_shared_definition_options():
    options = []

    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        return options

    rows = []
    for grp in def_file.Groups:
        for definition in grp.Definitions:
            rows.append((grp.Name or "", definition.Name or "", definition))

    rows.sort(key=lambda row: (row[0].lower(), row[1].lower()))
    for grp_name, def_name, definition in rows:
        label = "{} / {}".format(grp_name, def_name)
        options.append(OptionItem(label, definition))

    return options


def _is_yesno_parameter(fp):
    definition = getattr(fp, "Definition", None)
    if definition is None:
        return False

    try:
        ptype = str(definition.ParameterType)
        if ptype.lower() in ["yesno", "yes/no"]:
            return True
    except Exception:
        pass

    try:
        return "yes/no" in _parameter_data_type_label(fp).lower()
    except Exception:
        return False


def _try_parse_double(text, family_param):
    cleaned = (text or "").strip()
    if not cleaned:
        raise ValueError("Initial value is empty.")

    # Explicit unit suffix support for length-like inputs.
    # Examples: "3000mm", "3.2 m", "12in", "10 ft"
    lower = cleaned.lower()
    if lower.endswith("mm"):
        num = float(cleaned[:-2].strip().replace(",", "."))
        return num / 304.8
    if lower.endswith("cm"):
        num = float(cleaned[:-2].strip().replace(",", "."))
        return num / 30.48
    if lower.endswith(" m") or lower.endswith("m"):
        num = float(cleaned[:-1].strip().replace(",", "."))
        return num / 0.3048
    if lower.endswith("in") or lower.endswith('"'):
        num_txt = cleaned[:-2].strip() if lower.endswith("in") else cleaned[:-1].strip()
        num = float(num_txt.replace(",", "."))
        return num / 12.0
    if lower.endswith("ft") or lower.endswith("'"):
        num_txt = cleaned[:-2].strip() if lower.endswith("ft") else cleaned[:-1].strip()
        return float(num_txt.replace(",", "."))

    # Bare numeric value: interpret it using the parameter's display unit.
    # This avoids the previous behavior where plain numbers were treated as feet.
    if cleaned.count(",") == 1 and cleaned.count(".") == 0:
        cleaned = cleaned.replace(",", ".")
    value = float(cleaned)

    try:
        definition = getattr(family_param, "Definition", None)
        get_data_type = getattr(definition, "GetDataType", None)
        if callable(get_data_type):
            spec_id = get_data_type()
            if spec_id is not None:
                units = doc.GetUnits()
                fmt = units.GetFormatOptions(spec_id)
                get_unit_type_id = getattr(fmt, "GetUnitTypeId", None)
                if callable(get_unit_type_id):
                    unit_type_id = get_unit_type_id()
                    if unit_type_id is not None:
                        return UnitUtils.ConvertToInternalUnits(value, unit_type_id)
    except Exception:
        pass

    # Fallback if unit metadata isn't available: preserve previous behavior.
    return value


def _try_parse_int_or_bool(text, yesno):
    cleaned = (text or "").strip().lower()
    if yesno:
        if cleaned in ["1", "true", "yes", "y", "on"]:
            return 1
        if cleaned in ["0", "false", "no", "n", "off"]:
            return 0
        raise ValueError("Yes/No expects true/false, yes/no, on/off, or 1/0.")
    return int(cleaned)


def _resolve_element_id_value(raw_text, fp):
    text = (raw_text or "").strip()
    if not text:
        return True, ElementId.InvalidElementId, ""

    lower = text.lower()
    if lower in ["none", "invalid", "null"]:
        return True, ElementId.InvalidElementId, ""

    try:
        return True, ElementId(int(text)), ""
    except Exception:
        pass

    type_label = _parameter_data_type_label(fp).lower()

    if "material" in type_label:
        mats = FilteredElementCollector(doc).OfClass(Material).ToElements()
        for mat in mats:
            if (mat.Name or "").lower() == lower:
                return True, mat.Id, ""
        return False, None, "Material '{}' was not found.".format(text)

    types = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
    for typ in types:
        try:
            name = typ.Name or ""
        except Exception:
            name = ""
        if name.lower() == lower:
            return True, typ.Id, ""

    return False, None, "Could not resolve ElementId value '{}' by name or integer id.".format(text)


def _apply_initial_value(family_manager, family_param, raw_text):
    current_type = family_manager.CurrentType
    if current_type is None:
        return False, "Family has no current type, so initial value cannot be applied."

    storage = family_param.StorageType
    try:
        if storage == StorageType.Double:
            # Prefer Revit's native parser so plain values (e.g. "20") are
            # interpreted in the document's display units, like the Family Types dialog.
            set_value_string = getattr(family_manager, "SetValueString", None)
            if callable(set_value_string):
                try:
                    set_value_string(family_param, raw_text)
                    return True, ""
                except Exception:
                    # Fall back to explicit conversion path below.
                    pass

            family_manager.Set(family_param, _try_parse_double(raw_text, family_param))
            return True, ""

        if storage == StorageType.Integer:
            val = _try_parse_int_or_bool(raw_text, _is_yesno_parameter(family_param))
            family_manager.Set(family_param, val)
            return True, ""

        if storage == StorageType.String:
            family_manager.Set(family_param, raw_text)
            return True, ""

        if storage == StorageType.ElementId:
            ok, elem_id, reason = _resolve_element_id_value(raw_text, family_param)
            if not ok:
                return False, reason
            family_manager.Set(family_param, elem_id)
            return True, ""

        return False, "Unsupported storage type for initial value: {}".format(storage)
    except Exception as ex:
        return False, str(ex)


def _next_duplicate_name(base_name, existing_names):
    existing = {name for name in existing_names if name}
    index = 1
    while True:
        candidate = "{}-{}".format(base_name, index)
        if candidate not in existing:
            return candidate
        index += 1


def _duplicate_family_parameter(family_manager, source_param, new_name):
    group_value = source_param.Definition.ParameterGroup
    is_instance = bool(getattr(source_param, "IsInstance", False))
    data_type = _get_data_type(source_param.Definition)

    if data_type is None:
        raise Exception("Could not resolve source parameter data type.")

    # Shared parameters cannot be duplicated with a new name using the same external definition.
    # Create a non-shared family parameter with matching metadata instead.
    return family_manager.AddParameter(new_name, group_value, data_type, is_instance)


def _formula_references_parameter(formula_text, parameter_name):
    if not formula_text or not parameter_name:
        return False

    try:
        pattern = r"(?<![A-Za-z0-9_]){}(?![A-Za-z0-9_])".format(re.escape(parameter_name))
        return re.search(pattern, formula_text, flags=re.IGNORECASE) is not None
    except Exception:
        return parameter_name.lower() in formula_text.lower()


def _copy_current_parameter_value(family_manager, source_param, target_param):
    current_type = family_manager.CurrentType
    if current_type is None:
        return

    storage = source_param.StorageType

    if storage == StorageType.Double:
        family_manager.Set(target_param, current_type.AsDouble(source_param))
        return

    if storage == StorageType.Integer:
        family_manager.Set(target_param, current_type.AsInteger(source_param))
        return

    if storage == StorageType.String:
        val = current_type.AsString(source_param)
        family_manager.Set(target_param, val if val is not None else "")
        return

    if storage == StorageType.ElementId:
        family_manager.Set(target_param, current_type.AsElementId(source_param))


def _force_family_recalc(document):
    """Force parameter recalculation/regeneration after API-driven formula edits."""
    try:
        document.Regenerate()
    except Exception:
        pass

    # Some Revit versions expose a document-level explicit parameter eval.
    try:
        eval_all = getattr(document, "EvaluateAllParameterValues", None)
        if callable(eval_all):
            eval_all()
    except Exception:
        pass


def _force_family_type_cycle(family_manager):
    """Force family formulas to fully evaluate by cycling current type and returning."""
    try:
        current = family_manager.CurrentType
        if current is None:
            return

        all_types = []
        try:
            all_types = [t for t in family_manager.Types]
        except Exception:
            all_types = []

        if not all_types:
            family_manager.CurrentType = current
            return

        alt = None
        for t in all_types:
            try:
                if t.Id != current.Id:
                    alt = t
                    break
            except Exception:
                continue

        if alt is not None:
            family_manager.CurrentType = alt

        family_manager.CurrentType = current
    except Exception:
        pass


def _guard_or_exit():
    if not doc.IsFamilyDocument:
        forms.alert(
            "This tool only works in the Family Editor.\nOpen an .rfa file and run again.",
            exitscript=True,
        )
        return False

    return True


def main():
    if not _guard_or_exit():
        return

    fm = doc.FamilyManager
    if fm is None:
        forms.alert("FamilyManager is unavailable in this document.", exitscript=True)
        return

    dlg = ParameterEditorWindow(fm)
    dlg.ShowDialog()


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        forms.alert("Parameter Editor failed:\n{}".format(ex), title="Error")
        sys.exit()
