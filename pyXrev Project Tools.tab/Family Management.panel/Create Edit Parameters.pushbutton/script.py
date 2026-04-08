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
    ElementId,
    FilteredElementCollector,
    LabelUtils,
    Material,
    StorageType,
    Transaction,
    UnitUtils,
)

# BuiltInParameterGroup and ParameterType were removed in Revit 2025.
# Import them conditionally so the script loads on all versions 2020-2026.
try:
    from Autodesk.Revit.DB import BuiltInParameterGroup
    _BIPG_AVAILABLE = True
except ImportError:
    _BIPG_AVAILABLE = False

try:
    from Autodesk.Revit.DB import ParameterType
    _PT_AVAILABLE = True
except ImportError:
    _PT_AVAILABLE = False
from pyrevit import forms, revit
from formula_highlight import FormulaEditorHighlightMixin
from formula_analyzer import analyze_formula as _analyze_formula, replace_formula_subexpr as _replace_formula_subexpr
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


class ReorderItem(object):
    def __init__(self, fp):
        self.Param = fp
        self.Name = _param_name(fp)
        self.InstanceTypeLabel = "Instance" if bool(getattr(fp, "IsInstance", False)) else "Type"


class FormulaAnalysisWindow(forms.WPFWindow):
    """Comparison dialog showing original vs. simplified formula with colorized diff."""

    def __init__(self, original, result, on_apply_cb,
                 fm=None, current_item=None, all_items=None, autocomplete_names=None):
        forms.WPFWindow.__init__(self, "FormulaAnalysis.xaml")
        self._result = result
        self._on_apply_cb = on_apply_cb
        self._fm = fm
        self._current_item = current_item
        self._all_items = all_items or []
        self._autocomplete_names = set(autocomplete_names or [])

        # Import WPF inline types needed for colorized text
        from System.Windows.Documents import Run
        from System.Windows.Media import SolidColorBrush, Color

        def _color(r, g, b):
            return SolidColorBrush(Color.FromRgb(r, g, b))

        _grey  = _color(0x2A, 0x2A, 0x2A)
        _red   = _color(0xC6, 0x28, 0x28)
        _green = _color(0x2E, 0x7D, 0x32)

        def _populate(textblock, diff, highlight_kind, highlight_color):
            textblock.Inlines.Clear()
            space = Run(u" ")
            for kind, tok in diff:
                r = Run(tok)
                if kind == highlight_kind:
                    r.Foreground = highlight_color
                else:
                    r.Foreground = _grey
                textblock.Inlines.Add(r)
                textblock.Inlines.Add(Run(u" "))

        diff = result.get('diff', [])
        _populate(self.txtOriginalFormula,  diff, 'removed', _red)
        _populate(self.txtSuggestedFormula, diff, 'added',   _green)

        # Suggestions navigation
        self._suggestions = result.get('suggestions', [])
        self._suggestion_idx = 0
        self._show_suggestion(0)

        # Metrics
        m = result.get('metrics', {})
        if m:
            self.txtMetrics.Text = (
                u"Nodes: {}  |  Operator depth: {}  |  Unique parameters: {}".format(
                    m.get('nodes', '-'), m.get('depth', '-'), m.get('unique_params', '-')
                )
            )

        # Disable Apply when the simplified form is identical to the original
        simplified = result.get('simplified', original)
        self.btnApply.IsEnabled = (simplified != original)

    def _show_suggestion(self, idx):
        total = len(self._suggestions)
        if total == 0:
            self.txtSuggestionText.Text = u"No suggestions."
            self.txtSuggestionCounter.Text = u"0 / 0"
            self.btnPrevSuggestion.IsEnabled = False
            self.btnNextSuggestion.IsEnabled = False
            self.btnCopySuggestion.IsEnabled = False
            return
        sug = self._suggestions[idx]
        self.txtSuggestionCounter.Text = u"{} / {}".format(idx + 1, total)
        self.btnPrevSuggestion.IsEnabled = idx > 0
        self.btnNextSuggestion.IsEnabled = idx < total - 1
        self.btnCopySuggestion.IsEnabled = True
        if isinstance(sug, dict) and sug.get('type') == 'repeated_subexpr':
            self._render_repeated_subexpr(sug)
        else:
            self.pnlCreateReplace.Visibility = System.Windows.Visibility.Collapsed
            self.txtSuggestionText.Inlines.Clear()
            from System.Windows.Documents import Run
            self.txtSuggestionText.Inlines.Add(Run(sug if isinstance(sug, str) else str(sug)))

    def _render_repeated_subexpr(self, sug):
        from System.Windows.Documents import Run, LineBreak
        from System.Windows.Media import SolidColorBrush, Color, FontFamily
        from System.Windows import FontWeights, FontStyles

        def _col(r, g, b):
            return SolidColorBrush(Color.FromRgb(r, g, b))

        tb = self.txtSuggestionText
        tb.Inlines.Clear()

        # Header
        hdr = Run(u"Repeated subexpression")
        hdr.FontWeight = FontWeights.Bold
        tb.Inlines.Add(hdr)
        tb.Inlines.Add(LineBreak())

        # Subexpr in monospace blue
        expr = Run(u"  " + sug['subexpr'])
        expr.FontFamily = FontFamily("Consolas")
        expr.Foreground = _col(0x1A, 0x53, 0x7E)
        tb.Inlines.Add(expr)
        tb.Inlines.Add(LineBreak())
        tb.Inlines.Add(LineBreak())

        # Count in this formula
        count_self = sug.get('count_self', 0)
        tb.Inlines.Add(Run(u"  {} x in this parameter".format(count_self)))
        tb.Inlines.Add(LineBreak())

        # Whole matches
        for pname, cnt in sorted(sug.get('whole_matches', {}).items()):
            r = Run(u"  {} x as entire formula of parameter \"{}\"".format(cnt, pname))
            r.Foreground = _col(0x2E, 0x7D, 0x32)
            tb.Inlines.Add(r)
            tb.Inlines.Add(LineBreak())

        # Partial matches
        for pname, cnt in sorted(sug.get('partial_matches', {}).items()):
            r = Run(u"  {} x in formula of parameter \"{}\"".format(cnt, pname))
            r.Foreground = _col(0x1A, 0x53, 0x7E)
            tb.Inlines.Add(r)
            tb.Inlines.Add(LineBreak())

        # Advice
        tb.Inlines.Add(LineBreak())
        if sug.get('whole_matches'):
            advice = u"Consider replacing all occurrences with the parameter name directly."
        else:
            advice = u"Consider extracting it to a dedicated intermediate parameter."
        adv = Run(advice)
        adv.Foreground = _col(0x5A, 0x5A, 0x5A)
        adv.FontStyle = FontStyles.Italic
        tb.Inlines.Add(adv)

        # Populate and show Create & Replace panel
        self.txtNewParamName.Text = self._suggest_subexpr_name(sug['subexpr'])
        is_instance = self._infer_instance_scope(sug)
        self.rbNewParamInstance.IsChecked = is_instance
        self.rbNewParamType.IsChecked = not is_instance

        # Populate data type combo
        type_options = _build_type_options(self._fm) if self._fm else []
        self.cmbNewDataType.ItemsSource = type_options
        if type_options:
            self.cmbNewDataType.SelectedIndex = 0
        inferred_dt = _infer_subexpr_datatype(sug['subexpr'], self._all_items)
        if inferred_dt is not None:
            inferred_str = str(inferred_dt)
            for i, opt in enumerate(type_options):
                if str(opt.Value) == inferred_str:
                    self.cmbNewDataType.SelectedIndex = i
                    break

        self.pnlCreateReplace.Visibility = System.Windows.Visibility.Visible

    def _suggest_subexpr_name(self, subexpr_text):
        """Derive a parameter name suggestion from the subexpression text."""
        import re as _re
        existing = {it.Name.lower() for it in self._all_items if it.Name}
        _revit_fns = frozenset([
            'if', 'and', 'or', 'not', 'sqrt', 'round', 'roundup', 'rounddown',
            'abs', 'log', 'exp', 'sin', 'cos', 'tan',
        ])
        idents = [w for w in _re.findall(r'[A-Za-z_]\w*', subexpr_text)
                  if w.lower() not in _revit_fns]
        parts = idents[:2]
        base = ('_'.join(p[:15] for p in parts) + '_Calc') if parts else 'Calculated'
        name = base
        i = 1
        while name.lower() in existing:
            name = '{}_{}'.format(base, i)
            i += 1
        return name

    def _infer_instance_scope(self, sug):
        """Return True (Instance) if any affected parameter is Instance, else False (Type)."""
        if self._current_item and self._current_item.InstanceTypeLabel == 'Instance':
            return True
        affected_names = (list(sug.get('whole_matches', {}).keys()) +
                          list(sug.get('partial_matches', {}).keys()))
        for pname in affected_names:
            for it in self._all_items:
                if it.Name == pname and it.InstanceTypeLabel == 'Instance':
                    return True
        return False

    def on_create_and_replace(self, sender, args):
        sug = self._suggestions[self._suggestion_idx]
        if not isinstance(sug, dict) or sug.get('type') != 'repeated_subexpr':
            return

        new_name = (self.txtNewParamName.Text or u'').strip()
        if not new_name:
            forms.alert(u'Please enter a parameter name.', title='Create & Replace')
            return

        existing = {it.Name.lower() for it in self._all_items if it.Name}
        if new_name.lower() in existing:
            forms.alert(
                u"A parameter named '{}' already exists.".format(new_name),
                title='Create & Replace'
            )
            return

        rep_key = sug.get('rep_key')
        if not rep_key:
            forms.alert(u'Cannot determine subexpression key.', title='Create & Replace')
            return

        if self._current_item is None or self._fm is None:
            forms.alert(u'No parameter context available.', title='Create & Replace')
            return

        dt_item = self.cmbNewDataType.SelectedItem
        data_type = dt_item.Value if dt_item else None
        if data_type is None:
            forms.alert(u'Please select a data type for the new parameter.', title='Create & Replace')
            return

        is_instance = bool(self.rbNewParamInstance.IsChecked)
        param_names = self._autocomplete_names
        subexpr_text = sug['subexpr']

        # Build list of ParameterItems whose formulas will be rewritten
        affected = []
        if self._current_item and self._current_item.Formula:
            affected.append(self._current_item)
        for pname in (list(sug.get('whole_matches', {}).keys()) +
                      list(sug.get('partial_matches', {}).keys())):
            for it in self._all_items:
                if it.Name == pname and it.Formula and it not in affected:
                    affected.append(it)

        # Pre-compute new formulas (before opening the transaction)
        new_formulas = {}
        for it in affected:
            try:
                new_f = _replace_formula_subexpr(it.Formula, rep_key, new_name, param_names)
                new_formulas[it.Name] = new_f
            except Exception as ex:
                forms.alert(
                    u"Could not compute replacement formula for '{}':\n{}".format(it.Name, ex),
                    title='Create & Replace'
                )
                return

        # Dump diagnostics to a copyable pyRevit output window
        from pyrevit import script as _pyscript
        _out = _pyscript.get_output()
        _out.print_md(u'## Create & Replace — diagnostics')
        _out.print_md(u'**New parameter:** `{}`  \n**Subexpression formula:** `{}`'.format(
            new_name, subexpr_text))
        _out.print_md(u'### Rewritten formulas')
        for _n, _f in new_formulas.items():
            _out.print_md(u'**{}**  \n```\n{}\n```'.format(_n, _f))

        # Transaction 1: create the new parameter and set its formula.
        # Must be committed before Revit will accept it as a valid reference
        # in other parameters' formulas.
        txn = None
        try:
            txn = Transaction(doc, u'Create Subexpression Parameter')
            txn.Start()
            new_fp = self._fm.AddParameter(
                new_name, _default_group_type_general(), data_type, is_instance
            )
            self._fm.SetFormula(new_fp, subexpr_text)
            txn.Commit()
        except Exception as ex:
            try:
                if txn is not None and txn.GetStatus().ToString() == 'Started':
                    txn.RollBack()
            except Exception:
                pass
            _out.print_md(u'**ERROR (creating parameter):** {}'.format(ex))
            forms.alert(
                u'Create & Replace failed (creating parameter):\n{}'.format(ex),
                title='Create & Replace'
            )
            return

        # Transaction 2: rewrite affected parameters to reference the new parameter.
        txn2 = None
        try:
            txn2 = Transaction(doc, u'Replace Subexpression Occurrences')
            txn2.Start()
            fp_by_name = {_param_name(fp): fp for fp in _get_family_parameters(self._fm)}
            for it_name, new_f in new_formulas.items():
                fp = fp_by_name.get(it_name)
                if fp:
                    try:
                        self._fm.SetFormula(fp, new_f)
                        _out.print_md(u'OK: `{}`'.format(it_name))
                    except Exception as inner:
                        _out.print_md(
                            u'**FAILED:** `{}`  \n**Formula:** `{}`  \n**Error:** {}'.format(
                                it_name, new_f, inner)
                        )
                        raise Exception(
                            u"SetFormula failed on '{}'.\nFormula:\n{}\nError: {}".format(
                                it_name, new_f, inner)
                        )
                else:
                    _out.print_md(u'WARNING: parameter `{}` not found in family.'.format(it_name))
            txn2.Commit()
        except Exception as ex:
            try:
                if txn2 is not None and txn2.GetStatus().ToString() == 'Started':
                    txn2.RollBack()
            except Exception:
                pass
            forms.alert(
                u'Create & Replace failed (rewriting formulas):\n'
                u'See the pyRevit output window for full details.',
                title='Create & Replace'
            )
            return

        # Signal parent to reload without setting a specific formula text
        if self._on_apply_cb:
            self._on_apply_cb(None)
        self.Close()

    def on_prev_suggestion(self, sender, args):
        if self._suggestion_idx > 0:
            self._suggestion_idx -= 1
            self._show_suggestion(self._suggestion_idx)

    def on_next_suggestion(self, sender, args):
        if self._suggestion_idx < len(self._suggestions) - 1:
            self._suggestion_idx += 1
            self._show_suggestion(self._suggestion_idx)

    def on_copy_suggestion(self, sender, args):
        from System.Windows import Clipboard
        if not self._suggestions:
            return
        sug = self._suggestions[self._suggestion_idx]
        if isinstance(sug, dict) and sug.get('type') == 'repeated_subexpr':
            lines = [u"Repeated subexpression: " + sug['subexpr'],
                     u"  {} x in this parameter".format(sug.get('count_self', 0))]
            for pname, cnt in sorted(sug.get('whole_matches', {}).items()):
                lines.append(u"  {} x as entire formula of parameter \"{}\"".format(cnt, pname))
            for pname, cnt in sorted(sug.get('partial_matches', {}).items()):
                lines.append(u"  {} x in formula of parameter \"{}\"".format(cnt, pname))
            text = u"\n".join(lines)
        else:
            text = sug if isinstance(sug, str) else str(sug)
        Clipboard.SetText(text)

    def on_apply(self, sender, args):
        simplified = self._result.get('simplified', '')
        if simplified and self._on_apply_cb:
            self._on_apply_cb(simplified)
        self.Close()

    def on_close(self, sender, args):
        self.Close()


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


class ParameterEditorWindow(FormulaEditorHighlightMixin, forms.WPFWindow):
    def __init__(self, family_manager):
        forms.WPFWindow.__init__(self, "ParameterEditor.xaml")
        self.fm = family_manager
        self._all_groups_label = "All Groups"
        self._all_items = []
        self._filtered_items = ObservableCollection[object]()
        self._autocomplete_names = []
        self._shared_options = []
        self._active_token_range = None
        self._active_new_token_range = None
        self._highlighting = False

        self.lstParameters.ItemsSource = self._filtered_items

        self._all_types_by_discipline = _build_all_type_options_by_discipline()
        self._discipline_options = sorted(self._all_types_by_discipline.keys())
        self._group_options = _build_group_options(self.fm)
        self._shared_options = _build_shared_definition_options()

        self.cmbNewDiscipline.ItemsSource = self._discipline_options
        _default_disc = "Common" if "Common" in self._discipline_options else (
            self._discipline_options[0] if self._discipline_options else None
        )
        if _default_disc:
            self.cmbNewDiscipline.SelectedItem = _default_disc
        self._bind_combo_options(self.cmbNewType, self._type_options_for_discipline(_default_disc or ""))
        self._bind_combo_options(self.cmbNewGroup, self._group_options)
        self._bind_combo_options(self.cmbSharedDefinition, self._shared_options)
        self._bind_combo_options(self.cmbEditGroup, self._group_options)
        self._bind_combo_options(self.cmbBatchGroup, self._group_options)

        self._set_shared_mode(False)
        self._reload_parameter_items(select_name=None)
        self._sort_group_checkboxes = []
        self._populate_sort_groups()
        self._reorder_items = ObservableCollection[object]()
        self.lstReorderParams.ItemsSource = self._reorder_items
        self._populate_reorder_groups()
        self._restore_window_position()
        self._restore_column_layout()
        self._update_batch_panel()
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
        if hasattr(self, '_reorder_items'):
            self._populate_reorder_groups()

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
            if search:
                match_name = search in hay_name
                match_formula = bool(self.chkSearchFormula.IsChecked) and search in hay_formula
                if not match_name and not match_formula:
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
            self._formula_set_text("")
            self.txtFormulaBracketStatus.Text = ""
            self.btnToggleInstanceType.Content = ""
            self.btnToggleInstanceType.IsEnabled = False
            self.txtFormula.BorderBrush = self._brush("#B5B5B5")
            self.txtFormula.BorderThickness = self._uniform_thickness(1)
            return

        self.txtSelectedParamName.Text = item.Name
        self.txtSelectedParamDataType.Text = item.DataTypeLabel
        self.txtSelectedParamValue.Text = item.Value
        self.txtRenameTo.Text = item.Name
        self._formula_set_text(item.Formula or "")
        self._set_edit_group_selection(item.Param)
        self.btnToggleInstanceType.Content = item.InstanceTypeLabel
        self.btnToggleInstanceType.IsEnabled = not item.IsShared
        self._update_formula_bracket_feedback()

    def _set_edit_group_selection(self, family_param):
        try:
            current_group = _get_group_type(family_param.Definition)
        except Exception:
            current_group = None

        if current_group is None:
            return

        for opt in self._group_options:
            if str(opt.Value) == str(current_group):
                self.cmbEditGroup.SelectedItem = opt
                return

    # ── Multi-select / Batch helpers ──────────────────────────────────────

    def _selected_parameter_items(self):
        """Return all currently selected ParameterItems as a plain Python list."""
        return list(self.lstParameters.SelectedItems)

    def _update_batch_panel(self):
        """Refresh the Batch Edit tab controls to reflect the current selection."""
        items = self._selected_parameter_items()
        count = len(items)

        from System.Windows.Media import SolidColorBrush, Color

        if count == 0:
            self.tabBatchEdit.Background = SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0))
            self.tabBatchEdit.Foreground = SolidColorBrush(Color.FromRgb(0x00, 0x00, 0x00))
            self.txtBatchSummary.Text = (
                u"Select multiple parameters from the list on the left to batch edit them."
            )
            self.btnBatchToggle.IsEnabled = False
            self.btnBatchToggle.Content = u"Toggle"
            self.txtBatchToggleInfo.Text = u""
            self.btnBatchChangeGroup.IsEnabled = False
            self.btnBatchRename.IsEnabled = False
            self.btnBatchDuplicate.IsEnabled = False
            self.btnBatchDelete.IsEnabled = False
            self._set_batch_group_combo([])
            return

        if count >= 2:
            self.tabBatchEdit.Background = SolidColorBrush(Color.FromRgb(0x1E, 0x6D, 0xBB))
            if self.tabBatchEdit.IsSelected:
                self.tabBatchEdit.Foreground = SolidColorBrush(Color.FromRgb(0x00, 0x00, 0x00))
            else:
                self.tabBatchEdit.Foreground = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
        else:
            self.tabBatchEdit.Background = SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0))
            self.tabBatchEdit.Foreground = SolidColorBrush(Color.FromRgb(0x00, 0x00, 0x00))
        self.txtBatchSummary.Text = u"{} parameter(s) selected.".format(count)
        self.btnBatchChangeGroup.IsEnabled = True
        self.btnBatchRename.IsEnabled = True
        self.btnBatchDuplicate.IsEnabled = True
        self.btnBatchDelete.IsEnabled = True

        # Toggle: active only when all non-shared items share the same state
        non_shared = [it for it in items if not it.IsShared]
        if not non_shared:
            self.btnBatchToggle.IsEnabled = False
            self.btnBatchToggle.Content = u"Toggle"
            self.txtBatchToggleInfo.Text = u"All selected are shared (cannot toggle)."
        else:
            states = set(it.InstanceTypeLabel for it in non_shared)
            if len(states) == 1:
                current_state = list(states)[0]
                target_state = u"Type" if current_state == u"Instance" else u"Instance"
                self.btnBatchToggle.IsEnabled = True
                self.btnBatchToggle.Content = u"Make {}".format(target_state)
                info = u"All {} non-shared are {}.".format(len(non_shared), current_state)
                skipped = len(items) - len(non_shared)
                if skipped:
                    info += u"  ({} shared will be skipped.)".format(skipped)
                self.txtBatchToggleInfo.Text = info
            else:
                self.btnBatchToggle.IsEnabled = False
                self.btnBatchToggle.Content = u"Toggle"
                self.txtBatchToggleInfo.Text = (
                    u"Mixed Instance/Type \u2014 select all same state to enable toggle."
                )

        # Group combo: <Varies> when groups differ, pre-select when all same
        self._set_batch_group_combo(items)

    def _set_batch_group_combo(self, items):
        """Populate cmbBatchGroup; prepend <Varies> when selected params span groups."""
        if not items:
            self._bind_combo_options(self.cmbBatchGroup, self._group_options)
            return

        groups = set(it.GroupLabel for it in items)
        options = list(self._group_options)
        if len(groups) > 1:
            varies = OptionItem(u"<Varies>", None)
            options = [varies] + options
            self.cmbBatchGroup.ItemsSource = options
            self.cmbBatchGroup.DisplayMemberPath = u"Label"
            self.cmbBatchGroup.SelectedValuePath = u"Value"
            self.cmbBatchGroup.SelectedIndex = 0
        else:
            self._bind_combo_options(self.cmbBatchGroup, options)
            target_label = list(groups)[0]
            for opt in options:
                if opt.Label == target_label:
                    self.cmbBatchGroup.SelectedItem = opt
                    break

    def _do_group_move(self, fp, target_group):
        """Move fp to target_group using whatever API is available.  Raises on failure."""
        m = getattr(self.fm, "MoveParameter", None)
        if callable(m):
            m(fp, target_group)
            return

        m = getattr(self.fm, "SetParameterGroup", None)
        if callable(m):
            m(fp, target_group)
            return

        try:
            m = self.fm.GetType().GetMethod("SetParameterGroup")
            if m is not None:
                m.Invoke(self.fm, System.Array[System.Object]([fp, target_group]))
                return
        except Exception:
            pass

        try:
            defn = fp.Definition
            prop = defn.GetType().GetProperty("ParameterGroup")
            if prop is not None and prop.CanWrite:
                prop.SetValue(defn, target_group, None)
                return
        except Exception:
            pass

        raise Exception("Group move is not available in this Revit version.")

    def _set_shared_mode(self, is_shared):
        self.chkNewShared.IsChecked = is_shared
        self.cmbSharedDefinition.IsEnabled = is_shared
        self.cmbNewDiscipline.IsEnabled = not is_shared
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
        txt = self._formula_get_text()
        rng = self._selected_token_range(txt, self._formula_get_caret())
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

        text = self._formula_get_text()
        rng = self._active_token_range or self._selected_token_range(text, self._formula_get_caret())
        if rng is None:
            return False

        new_text = text[:rng[0]] + suggestion + text[rng[1]:]
        caret_pos = rng[0] + len(suggestion)
        self._formula_set_text(new_text)
        self.popupSuggestions.IsOpen = False

        self.txtFormula.Focus()
        self._formula_set_caret(caret_pos)
        return True

    # ── Create-tab formula helpers (parallel to edit-tab versions above) ─────

    def _render_new_suggestions(self):
        txt = self._rtb_get_text(self.txtNewFormula)
        rng = self._selected_token_range(txt, self._rtb_get_caret(self.txtNewFormula))
        self._active_new_token_range = rng

        if rng is None:
            self.popupNewSuggestions.IsOpen = False
            return

        token = txt[rng[0]:rng[1]].strip()
        if not token:
            self.popupNewSuggestions.IsOpen = False
            return

        token_lower = token.lower()
        prefix = []
        contains = []
        for name in self._autocomplete_names:
            if not name:
                continue
            low = name.lower()
            if low.startswith(token_lower):
                prefix.append(name)
            elif token_lower in low:
                contains.append(name)

        suggestions = (prefix + contains)[:40]
        self.lstNewSuggestions.ItemsSource = suggestions
        if suggestions:
            self.lstNewSuggestions.SelectedIndex = 0
            self.popupNewSuggestions.IsOpen = True
        else:
            self.popupNewSuggestions.IsOpen = False

    def _commit_new_suggestion(self):
        if not self.popupNewSuggestions.IsOpen:
            return False

        suggestion = self.lstNewSuggestions.SelectedItem
        if suggestion is None:
            return False

        text = self._rtb_get_text(self.txtNewFormula)
        rng = self._active_new_token_range or self._selected_token_range(text, self._rtb_get_caret(self.txtNewFormula))
        if rng is None:
            return False

        new_text = text[:rng[0]] + suggestion + text[rng[1]:]
        caret_pos = rng[0] + len(suggestion)
        self._rtb_set_text(self.txtNewFormula, new_text)
        self.popupNewSuggestions.IsOpen = False

        self.txtNewFormula.Focus()
        self._rtb_set_caret(self.txtNewFormula, caret_pos)
        return True

    def _set_new_formula_feedback(self, message, color_hex, border_hex=None, border_size=1):
        self.txtNewFormulaBracketStatus.Text = message or ""
        self.txtNewFormulaBracketStatus.Foreground = self._brush(color_hex)

        if border_hex is None:
            border_hex = "#B5B5B5"
        self.txtNewFormula.BorderBrush = self._brush(border_hex)

    def _update_new_formula_bracket_feedback(self):
        text = self._rtb_get_text(self.txtNewFormula)
        caret = self._rtb_get_caret(self.txtNewFormula)
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
                self._set_new_formula_feedback(
                    "Bracket pair '{}' matched (positions {} and {}).".format(ch, near_idx + 1, match_idx + 1),
                    color,
                    color,
                    2,
                )
                return
            self._set_new_formula_feedback(
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
            self._set_new_formula_feedback(msg, "#B00020", "#B00020", 2)
            return

        if not text:
            self._set_new_formula_feedback("", "#6A6A6A", "#B5B5B5", 1)
            return

        self._set_new_formula_feedback("Brackets are balanced.", "#6A6A6A", "#B5B5B5", 1)

    def _insert_pair(self, opener, closer):
        text = self._formula_get_text()
        start = self._formula_get_selection_start()
        length = self._formula_get_selection_length()

        if length > 0:
            selected = text[start:start + length]
            new_text = text[:start] + opener + selected + closer + text[start + length:]
            self._formula_set_text(new_text)
            self._formula_select(start + length + 2, 0)
            return

        new_text = text[:start] + opener + closer + text[start:]
        self._formula_set_text(new_text)
        self._formula_select(start + 1, 0)

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

    def _update_formula_bracket_feedback(self):
        text = self._formula_get_text()
        caret = self._formula_get_caret()
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

    def on_search_formula_toggled(self, sender, args):
        self._apply_parameter_filter()

    def on_group_filter_changed(self, sender, args):
        self._apply_parameter_filter()

    def on_param_selected(self, sender, args):
        self._set_editor_from_selected()
        self.popupSuggestions.IsOpen = False
        self._update_batch_panel()

    def on_formula_changed(self, sender, args):
        if self._highlighting:
            return
        self._apply_syntax_highlights_to(self.txtFormula)
        self._render_suggestions()
        self._update_formula_bracket_feedback()

    def on_formula_selection_changed(self, sender, args):
        if self._highlighting:
            return
        self._apply_syntax_highlights_to(self.txtFormula)
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

    # ── Create-tab formula event handlers ────────────────────────────────────

    def on_new_formula_changed(self, sender, args):
        if self._highlighting:
            return
        self._apply_syntax_highlights_to(self.txtNewFormula)
        self._render_new_suggestions()
        self._update_new_formula_bracket_feedback()

    def on_new_formula_selection_changed(self, sender, args):
        if self._highlighting:
            return
        self._apply_syntax_highlights_to(self.txtNewFormula)
        self._update_new_formula_bracket_feedback()

    def on_new_formula_keydown(self, sender, args):
        if args.Key == Key.Down and self.popupNewSuggestions.IsOpen:
            idx = self.lstNewSuggestions.SelectedIndex
            if idx < self.lstNewSuggestions.Items.Count - 1:
                self.lstNewSuggestions.SelectedIndex = idx + 1
                self.lstNewSuggestions.ScrollIntoView(self.lstNewSuggestions.SelectedItem)
            args.Handled = True
            return

        if args.Key == Key.Up and self.popupNewSuggestions.IsOpen:
            idx = self.lstNewSuggestions.SelectedIndex
            if idx > 0:
                self.lstNewSuggestions.SelectedIndex = idx - 1
                self.lstNewSuggestions.ScrollIntoView(self.lstNewSuggestions.SelectedItem)
            args.Handled = True
            return

        if args.Key == Key.Enter and self.popupNewSuggestions.IsOpen:
            if self._commit_new_suggestion():
                args.Handled = True
            return

        if args.Key == Key.Tab and self.popupNewSuggestions.IsOpen:
            if self._commit_new_suggestion():
                args.Handled = True
            return

        if args.Key == Key.Escape and self.popupNewSuggestions.IsOpen:
            self.popupNewSuggestions.IsOpen = False
            args.Handled = True

    def on_new_suggestion_double_click(self, sender, args):
        self._commit_new_suggestion()

    def on_clear_new_formula(self, sender, args):
        self._rtb_set_text(self.txtNewFormula, "")

    def _perform_toggle_single(self, item):
        """
        Perform the full Instance/Type toggle for a single ParameterItem, including
        all chain-dependency checks and user dialogs.

        Returns ('ok', message), ('cancelled', message), or ('error', message).
        Does NOT call _reload_parameter_items or _set_status — that is the caller's
        responsibility.
        """
        if item.IsShared:
            return ('error', u"Shared parameters cannot have their instance/type changed.")

        source_param = item.Param
        is_currently_instance = bool(getattr(source_param, "IsInstance", False))

        # ── Instance → Type: check if the formula references Instance params ──
        if is_currently_instance:
            chain, children_of = _build_instance_to_type_chain(self.fm, item.Name)
            if len(chain) > 1:
                dependent_names = []
                for level in chain[1:]:
                    dependent_names.extend(level)

                tree_text = _format_chain_tree(item.Name, children_of)
                num_levels = len(chain) - 1
                msg = (
                    u"Converting '{}' from Instance to Type requires converting {} dependent "
                    u"Instance parameter(s) across {} level(s) first.\n\n"
                    u"{}\n"
                    u"Conversion order: deepest level first, then '{}' last."
                ).format(
                    item.Name,
                    len(dependent_names),
                    num_levels,
                    tree_text,
                    item.Name,
                )

                choice = forms.alert(
                    msg,
                    title=u"Instance-to-Type: {} dependent param(s), {} level(s)".format(
                        len(dependent_names), num_levels
                    ),
                    options=["Convert All", "Cancel"],
                )

                if choice != "Convert All":
                    return ('cancelled', u"Toggle cancelled for '{}'.".format(item.Name))

                # Convert deepest level first, then shallower, finally the target.
                ordered_names = []
                for level in reversed(chain[1:]):
                    ordered_names.extend(level)
                ordered_names.append(item.Name)

                # Each MakeType must be its own committed transaction so Revit
                # sees the updated scope before validating the next parameter's formula.
                for name in ordered_names:
                    try:
                        with revit.Transaction(u"Convert '{}' to Type".format(name)):
                            live_fps = {_param_name(fp): fp for fp in _get_family_parameters(self.fm)}
                            fp = live_fps.get(name)
                            if fp is not None:
                                self.fm.MakeType(fp)
                    except Exception as ex:
                        return ('error', u"Failed converting '{}' to Type: {}".format(name, ex))

                converted = u", ".join(u"'{}'".format(n) for n in ordered_names)
                return ('ok', u"Converted to Type: {}.".format(converted))

        # ── Type → Instance: check for dependent Type parameters ─────────────
        if not is_currently_instance:
            chain, children_of = _build_type_to_instance_chain(self.fm, item.Name)
            if len(chain) > 1:
                dependent_names = []
                for level in chain[1:]:
                    dependent_names.extend(level)

                tree_text = _format_chain_tree(item.Name, children_of)
                num_levels = len(chain) - 1
                msg = (
                    u"Converting '{}' from Type to Instance requires converting {} dependent "
                    u"Type parameter(s) across {} level(s) first.\n\n"
                    u"{}\n"
                    u"Conversion order: deepest level first, then '{}' last."
                ).format(
                    item.Name,
                    len(dependent_names),
                    num_levels,
                    tree_text,
                    item.Name,
                )

                choice = forms.alert(
                    msg,
                    title=u"Type-to-Instance: {} dependent param(s), {} level(s)".format(
                        len(dependent_names), num_levels
                    ),
                    options=["Convert All", "Cancel"],
                )

                if choice != "Convert All":
                    return ('cancelled', u"Toggle cancelled for '{}'.".format(item.Name))

                # Convert deepest level first, then shallower, finally the target.
                ordered_names = []
                for level in reversed(chain[1:]):
                    ordered_names.extend(level)
                ordered_names.append(item.Name)

                fp_by_name = {_param_name(fp): fp for fp in _get_family_parameters(self.fm)}

                try:
                    with revit.Transaction("Convert Parameters to Instance"):
                        for name in ordered_names:
                            fp = fp_by_name.get(name)
                            if fp is not None:
                                self.fm.MakeInstance(fp)
                except Exception as ex:
                    return ('error', u"Convert all failed: {}".format(ex))

                converted = u", ".join(u"'{}'".format(n) for n in ordered_names)
                return ('ok', u"Converted to Instance: {}.".format(converted))

        # ── Normal single toggle ──────────────────────────────────────────────
        try:
            with revit.Transaction("Change Parameter Scope"):
                if is_currently_instance:
                    self.fm.MakeType(source_param)
                else:
                    self.fm.MakeInstance(source_param)
        except Exception as ex:
            return ('error', u"Toggle failed: {}".format(ex))

        new_scope = u"Type" if is_currently_instance else u"Instance"
        return ('ok', u"Changed '{}' to {}.".format(item.Name, new_scope))

    def _toggle_multiple(self, items):
        """
        Toggle Instance/Type for a batch of ParameterItems without pre-analysis dialogs.

        Instance→Type: uses retry passes so that chained params within the selection
        (e.g. param2=param1, param3=param2) all succeed — param1 converts first on
        pass 1, param2 on pass 2, etc.  Params that still fail after no-progress is
        detected (external formula mismatch) are reported clearly.

        Type→Instance: single pass in any order — Instance params may freely reference
        Type params so ordering never causes a failure.
        """
        non_shared = [it for it in items if not it.IsShared]
        shared_names = [it.Name for it in items if it.IsShared]

        if not non_shared:
            self._set_status(u"No non-shared parameters to toggle.", "error")
            return

        is_currently_instance = bool(getattr(non_shared[0].Param, "IsInstance", False))
        direction = u"Type" if is_currently_instance else u"Instance"

        ok_names = []
        error_names = []

        if is_currently_instance:
            # Instance → Type: must convert dependencies before dependents.
            # Pre-sort the selected items topologically so that a param whose formula
            # references another selected param comes AFTER that param in the list.
            names_in_set = {it.Name for it in non_shared}
            fp_snap = {_param_name(fp): fp for fp in _get_family_parameters(self.fm)}

            # Build internal dependency map: which selected params does each formula reference?
            internal_deps = {}
            for it in non_shared:
                fp = fp_snap.get(it.Name)
                formula = _safe_formula(fp) if fp else None
                if formula:
                    internal_deps[it.Name] = [
                        n for n in names_in_set
                        if n != it.Name and _formula_references_parameter(formula, n)
                    ]
                else:
                    internal_deps[it.Name] = []

            # Kahn's topological sort: items with no pending deps go first
            sorted_items = []
            remaining = list(non_shared)
            done_names = set()
            while remaining:
                progress = False
                next_remaining = []
                for it in remaining:
                    if all(d in done_names for d in internal_deps.get(it.Name, [])):
                        sorted_items.append(it)
                        done_names.add(it.Name)
                        progress = True
                    else:
                        next_remaining.append(it)
                if not progress:
                    sorted_items.extend(next_remaining)
                    break
                remaining = next_remaining

            # Convert in sorted order — each in its own transaction
            for it in sorted_items:
                try:
                    with revit.Transaction(u"Convert '{}' to Type".format(it.Name)):
                        live_fps = {_param_name(fp): fp for fp in _get_family_parameters(self.fm)}
                        fp = live_fps.get(it.Name)
                        if fp is not None:
                            self.fm.MakeType(fp)
                    ok_names.append(it.Name)
                except Exception:
                    error_names.append(it.Name)
        else:
            # Type → Instance: ordering matters in the SAME way but reversed.
            # A Type param whose formula references Param1 must be converted to
            # Instance BEFORE Param1 is converted — otherwise the Type param would
            # temporarily have an Instance param in its formula.
            # So we need REVERSE topological order: dependents (leaves) first.
            names_in_set = {it.Name for it in non_shared}
            fp_snap = {_param_name(fp): fp for fp in _get_family_parameters(self.fm)}

            # Build dependency map: which selected params does each formula reference?
            internal_deps = {}
            for it in non_shared:
                fp = fp_snap.get(it.Name)
                formula = _safe_formula(fp) if fp else None
                if formula:
                    internal_deps[it.Name] = [
                        n for n in names_in_set
                        if n != it.Name and _formula_references_parameter(formula, n)
                    ]
                else:
                    internal_deps[it.Name] = []

            # Topo sort (dependencies first), then reverse = dependents first
            topo = []
            remaining = list(non_shared)
            done_names = set()
            while remaining:
                progress = False
                next_remaining = []
                for it in remaining:
                    if all(d in done_names for d in internal_deps.get(it.Name, [])):
                        topo.append(it)
                        done_names.add(it.Name)
                        progress = True
                    else:
                        next_remaining.append(it)
                if not progress:
                    topo.extend(next_remaining)
                    break
                remaining = next_remaining

            # Reverse: convert leaves (dependents) first so no Type param ever
            # temporarily holds a formula referencing a newly-Instance param.
            sorted_items = list(reversed(topo))

            for it in sorted_items:
                try:
                    with revit.Transaction(u"Convert '{}' to Instance".format(it.Name)):
                        live_fps = {_param_name(fp): fp for fp in _get_family_parameters(self.fm)}
                        fp = live_fps.get(it.Name)
                        if fp is not None:
                            self.fm.MakeInstance(fp)
                    ok_names.append(it.Name)
                except Exception:
                    error_names.append(it.Name)

        parts = []
        if ok_names:
            parts.append(u"Converted {} to {}: {}.".format(
                len(ok_names), direction,
                u", ".join(u"'{}'".format(n) for n in ok_names)
            ))
        if shared_names:
            parts.append(u"{} shared parameter(s) skipped.".format(len(shared_names)))
        if error_names:
            parts.append(
                u"Could not convert to {}: {}.  "
                u"Each may be referenced in a formula with a Type/Instance mismatch.".format(
                    direction,
                    u", ".join(u"'{}'".format(n) for n in error_names)
                )
            )

        tone = "error" if (error_names and not ok_names) else ("ok" if ok_names else "neutral")
        self._set_status(u"  ".join(parts) if parts else u"Nothing to toggle.", tone)
        self._reload_parameter_items(select_name=None)
        self._populate_sort_groups()
        self._update_batch_panel()

    def on_toggle_instance_type(self, sender, args):
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        status, msg = self._perform_toggle_single(item)
        tone = "ok" if status == "ok" else ("neutral" if status == "cancelled" else "error")
        self._set_status(msg, tone)
        self._reload_parameter_items(select_name=item.Name)

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
        source_group = _get_group_type(source_param.Definition)

        if str(source_group) == str(target_group):
            self._set_status("Parameter is already in that group.", "neutral")
            return

        try:
            with revit.Transaction("Move Parameter Group"):
                self._do_group_move(source_param, target_group)
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

        formula_text = self._formula_get_text().strip()
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
        self._formula_set_text("")

    def on_analyze_formula(self, sender, args):
        formula = self._formula_get_text().strip()
        if not formula:
            self._set_status("No formula to analyze.", "neutral")
            return
        item = self._selected_parameter_item()
        if item is None:
            self._set_status("Select a parameter first.", "error")
            return

        # ── Step 1: apply to Revit so it normalises spacing/formatting ──
        txn = None
        try:
            txn = Transaction(doc, "Set Parameter Formula")
            txn.Start()
            self.fm.SetFormula(item.Param, formula)
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
            self._set_status("Formula apply failed \u2014 cannot analyse: {}".format(ex), "error")
            return

        # ── Step 2: reload and read back Revit's normalised formula text ──
        self._reload_parameter_items(select_name=item.Name)
        reloaded = self._selected_parameter_item()
        normalized = (reloaded.Formula if reloaded else None) or formula

        # ── Step 3: analyse the normalised formula ──
        current_name = reloaded.Name if reloaded else item.Name
        param_formulas = {
            it.Name: it.Formula
            for it in self._all_items
            if it.Name != current_name and it.Formula
        }
        result = _analyze_formula(
            normalized,
            param_names=set(self._autocomplete_names),
            param_formulas=param_formulas,
        )
        if not result.get('ok'):
            forms.alert(
                u"Formula analysis error:\n{}".format(result.get('error', 'Unknown error')),
                title="Formula Analysis",
            )
            return
        self._set_status("Formula applied and analysed.", "ok")
        win = FormulaAnalysisWindow(
            normalized, result, self.on_apply_formula_text,
            fm=self.fm,
            current_item=reloaded,
            all_items=self._all_items,
            autocomplete_names=self._autocomplete_names,
        )
        win.ShowDialog()

    def on_apply_formula_text(self, new_formula):
        if new_formula is not None:
            self._formula_set_text(new_formula)
            self.on_apply_formula(None, None)
        else:
            # Triggered by Create & Replace — just reload
            sel = self._selected_parameter_item()
            self._reload_parameter_items(select_name=sel.Name if sel else None)

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

    def _type_options_for_discipline(self, disc_str):
        """Return type OptionItems for disc_str, always including Common types."""
        disc_types = self._all_types_by_discipline.get(disc_str, [])
        if disc_str == 'Common':
            return disc_types
        common_types = self._all_types_by_discipline.get('Common', [])
        seen = set()
        merged = []
        for opt in common_types + disc_types:
            key = str(opt.Value)
            if key not in seen:
                seen.add(key)
                merged.append(opt)
        merged.sort(key=lambda o: o.Label.lower())
        return merged

    def on_new_discipline_changed(self, sender, args):
        disc = self.cmbNewDiscipline.SelectedItem
        if disc is None:
            return
        type_options = self._type_options_for_discipline(str(disc))
        self._bind_combo_options(self.cmbNewType, type_options)

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
        initial_formula_text = self._rtb_get_text(self.txtNewFormula).strip()

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
        self._rtb_set_text(self.txtNewFormula, "")
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

    # ── Reorder Parameters tab ────────────────────────────────────────────

    def _populate_reorder_groups(self):
        """Rebuild the group dropdown in the Reorder Parameters tab."""
        if not _sort_supports_reorder(self.fm):
            self.btnReorderApply.IsEnabled = False
            self.txtReorderApiWarning.Text = (
                u"FamilyManager.ReorderParameters is not available in this Revit version."
            )
            self.txtReorderApiWarning.Visibility = System.Windows.Visibility.Visible
            return
        self.btnReorderApply.IsEnabled = True
        self.txtReorderApiWarning.Visibility = System.Windows.Visibility.Collapsed

        all_params = _sort_get_current_order(self.fm)
        if not all_params:
            return

        grouped, labels = _sort_group_parameters(all_params)
        options = sorted(
            [OptionItem(labels[k], k) for k in grouped],
            key=lambda o: o.Label.lower()
        )
        self.cmbReorderGroup.ItemsSource = options
        self.cmbReorderGroup.DisplayMemberPath = "Label"
        self.cmbReorderGroup.SelectedValuePath = "Value"

        # Default to Dimensions; fall back to first item
        target = next((o for o in options if o.Label.lower() == "dimensions"), None)
        if target is None and options:
            target = options[0]
        if target is not None:
            self.cmbReorderGroup.SelectedItem = target
        self._populate_reorder_list()

    def _populate_reorder_list(self):
        """Fill lstReorderParams with the current group's params in Revit order."""
        self._reorder_items.Clear()
        selected_option = self.cmbReorderGroup.SelectedItem
        if selected_option is None:
            return
        group_key = selected_option.Value

        all_params = _sort_get_current_order(self.fm)
        if not all_params:
            return

        grouped, _ = _sort_group_parameters(all_params)
        for fp in grouped.get(group_key, []):
            self._reorder_items.Add(ReorderItem(fp))

    def on_reorder_group_changed(self, sender, args):
        self._populate_reorder_list()

    def on_reorder_refresh(self, sender, args):
        self._populate_reorder_groups()

    def _reorder_move(self, action):
        count = self._reorder_items.Count
        if count < 2:
            return

        # Identify selected items by identity so they survive a Clear+re-Add
        selected_ids = set(id(it) for it in self.lstReorderParams.SelectedItems)
        if not selected_ids:
            return

        items = list(self._reorder_items)
        selected_indices = sorted(i for i, it in enumerate(items) if id(it) in selected_ids)

        if action == 'top':
            shift = -selected_indices[0]
        elif action == 'up':
            shift = -1 if selected_indices[0] > 0 else 0
        elif action == 'down':
            shift = 1 if selected_indices[-1] < count - 1 else 0
        else:  # bottom
            shift = (count - 1) - selected_indices[-1]

        if shift == 0:
            return

        sel_set = set(selected_indices)
        selected_items = [items[i] for i in selected_indices]
        remaining = [items[i] for i in range(count) if i not in sel_set]

        if shift < 0:
            insert_at = max(0, selected_indices[0] + shift)
        else:
            # Number of non-selected items before the first selected item
            prefix_count = sum(1 for i in range(selected_indices[0]) if i not in sel_set)
            insert_at = min(len(remaining), prefix_count + shift)

        new_items = remaining[:insert_at] + selected_items + remaining[insert_at:]

        self._reorder_items.Clear()
        for it in new_items:
            self._reorder_items.Add(it)

        # Restore selection on the moved block
        self.lstReorderParams.SelectedItems.Clear()
        for i in range(insert_at, insert_at + len(selected_items)):
            self.lstReorderParams.SelectedItems.Add(self._reorder_items[i])
        self.lstReorderParams.ScrollIntoView(self._reorder_items[insert_at])

    def on_reorder_top(self, sender, args):
        self._reorder_move("top")

    def on_reorder_up(self, sender, args):
        self._reorder_move("up")

    def on_reorder_down(self, sender, args):
        self._reorder_move("down")

    def on_reorder_bottom(self, sender, args):
        self._reorder_move("bottom")

    def on_reorder_apply(self, sender, args):
        selected_option = self.cmbReorderGroup.SelectedItem
        if selected_option is None:
            self._set_status(u"Select a group to reorder.", "error")
            return
        group_key = selected_option.Value

        if self._reorder_items.Count == 0:
            self._set_status(u"No parameters in the selected group.", "neutral")
            return

        all_params = _sort_get_current_order(self.fm)
        if all_params is None:
            self._set_status(u"GetParameters API unavailable.", "error")
            return

        grouped, _ = _sort_group_parameters(all_params)
        group_names_new = {ri.Name for ri in self._reorder_items}

        # Build the new full order: replace the selected group's slice with
        # the user's arranged order; all other groups keep their positions.
        new_order_fps = []
        reorder_iter = iter(list(self._reorder_items))
        for fp in all_params:
            from sort_param_utils import get_group_info as _get_group_info
            key, _ = _get_group_info(fp)
            if key == group_key:
                try:
                    new_order_fps.append(next(reorder_iter).Param)
                except StopIteration:
                    pass  # shouldn't happen
            else:
                new_order_fps.append(fp)

        from System.Collections.Generic import List as _DotNetList
        from Autodesk.Revit.DB import FamilyParameter as _FamilyParameter
        dotnet_list = _DotNetList[_FamilyParameter]()
        for fp in new_order_fps:
            dotnet_list.Add(fp)

        try:
            with revit.Transaction(u"Reorder Parameters"):
                self.fm.ReorderParameters(dotnet_list)
        except Exception as ex:
            self._set_status(u"Reorder failed: {}".format(ex), "error")
            return

        self._set_status(
            u"Reordered '{}' group ({} parameter(s)).".format(
                selected_option.Label, self._reorder_items.Count
            ), "ok"
        )
        self._populate_reorder_list()  # refresh from new Revit order to confirm

    # ── Batch Edit tab handlers ───────────────────────────────────────────

    def on_tab_selection_changed(self, sender, args):
        # SelectionChanged bubbles up from ALL child controls (ComboBox, ListBox, etc.)
        # Only respond when an actual TabItem was added to the selection.
        from System.Windows.Controls import TabItem
        if not args.AddedItems or not isinstance(args.AddedItems[0], TabItem):
            return
        self._update_batch_panel()
        try:
            if self.tabImport.IsSelected:
                self._ensure_import_lib()
                self._populate_import_source_combo()
        except Exception:
            pass

    def on_batch_toggle(self, sender, args):
        items = self._selected_parameter_items()
        non_shared = [it for it in items if not it.IsShared]

        if not non_shared:
            self._set_status(u"No non-shared parameters selected.", "error")
            return
        if len(non_shared) == 1:
            status, msg = self._perform_toggle_single(non_shared[0])
            tone = "ok" if status == "ok" else ("neutral" if status == "cancelled" else "error")
            self._set_status(msg, tone)
            self._reload_parameter_items(select_name=non_shared[0].Name)
        else:
            self._toggle_multiple(items)
        self._update_batch_panel()

    def on_batch_change_group(self, sender, args):
        items = self._selected_parameter_items()
        if not items:
            self._set_status(u"Select parameters first.", "error")
            return

        selected_opt = self.cmbBatchGroup.SelectedItem
        if selected_opt is None or selected_opt.Value is None:
            self._set_status(u"Select a destination group.", "error")
            return

        target_group = selected_opt.Value
        target_label = _label_for_group(target_group)

        moved = []
        skipped = []
        errors = []
        for it in items:
            fp = it.Param
            try:
                source_group = _get_group_type(fp.Definition)
            except Exception:
                skipped.append(it.Name)
                continue

            if str(source_group) == str(target_group):
                skipped.append(it.Name)
                continue

            try:
                with revit.Transaction(u"Move Parameter Group"):
                    self._do_group_move(fp, target_group)
                moved.append(it.Name)
            except Exception as ex:
                errors.append(u"{} ({})".format(it.Name, ex))

        parts = []
        if moved:
            parts.append(u"Moved {} parameter(s) to '{}'.".format(len(moved), target_label))
        if skipped:
            parts.append(u"{} already in group (skipped).".format(len(skipped)))
        if errors:
            parts.append(u"Errors: {}.".format(u"  ".join(errors)))

        tone = "error" if errors else "ok"
        self._set_status(u"  ".join(parts) if parts else u"Nothing to move.", tone if parts else "neutral")
        self._reload_parameter_items(select_name=None)
        self._update_batch_panel()

    def on_batch_prefix_suffix(self, sender, args):
        items = self._selected_parameter_items()
        if not items:
            self._set_status(u"Select parameters first.", "error")
            return

        find_text    = (self.txtBatchFind.Text    or u"")
        replace_text = self.txtBatchReplace.Text if self.txtBatchReplace.Text is not None else u""
        prefix       = (self.txtBatchPrefix.Text  or u"").strip()
        suffix       = (self.txtBatchSuffix.Text  or u"").strip()
        use_regex    = bool(self.chkBatchRegex.IsChecked)

        if not find_text and not prefix and not suffix:
            self._set_status(u"Enter a Find value and/or Prefix/Suffix.", "error")
            return

        if find_text and use_regex:
            try:
                find_pattern = re.compile(find_text)
            except Exception as ex:
                self._set_status(u"Invalid regex: {}".format(ex), "error")
                return
        else:
            find_pattern = None

        rename_method = getattr(self.fm, "RenameParameter", None)
        if not callable(rename_method):
            self._set_status(u"Rename not supported in this Revit version.", "error")
            return

        # Shared params cannot be renamed — warn and offer to skip them
        shared_items = [it for it in items if it.IsShared]
        if shared_items:
            choice = forms.alert(
                u"Shared parameters cannot be renamed.\n\n"
                u"Skip {} shared parameter(s) and rename the rest?".format(len(shared_items)),
                title=u"Rename Shared Parameters",
                options=["Skip & Continue", "Cancel"],
            )
            if choice != "Skip & Continue":
                self._set_status(u"Rename cancelled.", "neutral")
                return
            items = [it for it in items if not it.IsShared]
            if not items:
                self._set_status(u"No renameable parameters selected.", "neutral")
                return

        def _apply_rename(name):
            # Apply find / replace on the original name first, then wrap with prefix / suffix
            result = name
            if find_text:
                if use_regex:
                    result = find_pattern.sub(replace_text, result)
                else:
                    result = result.replace(find_text, replace_text)
            return prefix + result + suffix

        # Pre-validate: compute all new names and check for conflicts
        existing_lower = {it.Name.lower() for it in self._all_items}
        selected_lower = {it.Name.lower() for it in items}
        available = existing_lower - selected_lower

        new_names = {}
        no_change = []
        conflicts = []
        for it in items:
            new_name = _apply_rename(it.Name)
            if not new_name.strip():
                conflicts.append(u"'{}' \u2192 (empty) \u2014 would produce an empty name".format(it.Name))
                continue
            if new_name == it.Name:
                no_change.append(it.Name)
                continue
            new_lower = new_name.lower()
            if new_lower in available or new_lower in {v.lower() for v in new_names.values()}:
                conflicts.append(u"'{}' \u2192 '{}' (already exists)".format(it.Name, new_name))
            else:
                new_names[it.Name] = new_name

        if conflicts:
            forms.alert(
                u"The following renames would create name conflicts:\n\n{}".format(
                    u"\n".join(conflicts)
                ),
                title=u"Rename Conflict",
            )
            return

        if not new_names:
            self._set_status(u"Find text not found in any selected parameter name.", "neutral")
            return

        renamed = []
        errors = []
        for it in items:
            new_name = new_names.get(it.Name)
            if new_name is None:
                continue
            try:
                with revit.Transaction(u"Rename Parameter"):
                    rename_method(it.Param, new_name)
                renamed.append(u"'{}' \u2192 '{}'".format(it.Name, new_name))
            except Exception as ex:
                errors.append(u"{} ({})".format(it.Name, ex))

        parts = []
        if renamed:
            parts.append(u"Renamed {} parameter(s).".format(len(renamed)))
        if no_change:
            parts.append(u"{} unchanged (no match).".format(len(no_change)))
        if errors:
            parts.append(u"Errors: {}.".format(u"  ".join(errors)))

        tone = "error" if errors else "ok"
        self._set_status(u"  ".join(parts) if parts else u"Nothing renamed.", tone if parts else "neutral")
        self._reload_parameter_items(select_name=None)
        self._populate_sort_groups()
        self._update_batch_panel()

    def on_batch_duplicate(self, sender, args):
        items = self._selected_parameter_items()
        if not items:
            self._set_status(u"Select parameters first.", "error")
            return

        shared_items = [it for it in items if it.IsShared]
        if shared_items:
            choice = forms.alert(
                u"Shared parameters cannot be duplicated with a new name using the same shared definition.\n"
                u"{} shared parameter(s) will be created as non-shared family parameters instead.\n\n"
                u"Continue?".format(len(shared_items)),
                title=u"Duplicate Shared Parameters",
                options=["Continue", "Cancel"],
            )
            if choice != "Continue":
                self._set_status(u"Duplicate cancelled.", "neutral")
                return

        # Track all names (existing + newly created) to avoid collisions within the batch
        all_names = [it.Name for it in self._all_items]
        created = []
        errors = []
        for it in items:
            new_name = _next_duplicate_name(it.Name, all_names)
            try:
                with revit.Transaction(u"Duplicate Family Parameter"):
                    new_param = _duplicate_family_parameter(self.fm, it.Param, new_name)
                    if not _safe_formula(it.Param):
                        _copy_current_parameter_value(self.fm, it.Param, new_param)
                    source_formula = _safe_formula(it.Param)
                    if source_formula:
                        self.fm.SetFormula(new_param, source_formula)
                all_names.append(new_name)
                created.append(new_name)
            except Exception as ex:
                errors.append(u"{} ({})".format(it.Name, ex))

        parts = []
        if created:
            parts.append(u"Duplicated {} parameter(s).".format(len(created)))
        if errors:
            parts.append(u"Errors: {}.".format(u"  ".join(errors)))

        tone = "error" if errors else "ok"
        self._set_status(u"  ".join(parts) if parts else u"Nothing duplicated.", tone if parts else "neutral")
        self._reload_parameter_items(select_name=None)
        self._populate_sort_groups()
        self._update_batch_panel()

    def on_batch_delete(self, sender, args):
        items = self._selected_parameter_items()
        if not items:
            self._set_status(u"Select parameters first.", "error")
            return

        names = [it.Name for it in items]

        try:
            used_set = find_directly_used_params(doc, self.fm)
            used_names_in_use = {_param_name(fp) for fp in used_set}
        except Exception:
            used_names_in_use = set()

        warning_lines = [
            u"Delete {} parameter(s)?\n\n{}\n\n"
            u"To undo, close the Parameter Editor and use Revit's Undo (Ctrl+Z).".format(
                len(names),
                u", ".join(u"'{}'".format(n) for n in names)
            )
        ]

        in_use = [n for n in names if n in used_names_in_use]
        if in_use:
            warning_lines.append(
                u"\u26a0  In use (dimension/association/array): {}.".format(
                    u", ".join(u"'{}'".format(n) for n in in_use)
                )
            )

        # Warn about formula references that survive the batch (not themselves being deleted)
        broken_refs = []
        for name in names:
            refs = find_formula_referencing_params(self.fm, name)
            external_refs = [r for r in refs if r not in names]
            if external_refs:
                broken_refs.append(
                    u"'{}' referenced by: {}.".format(name, u", ".join(external_refs))
                )
        if broken_refs:
            warning_lines.append(
                u"\u26a0  Formula references that will break:\n    " +
                u"\n    ".join(broken_refs)
            )

        confirmed = forms.alert(
            u"\n\n".join(warning_lines),
            title=u"Delete {} Parameters".format(len(names)),
            ok=True,
            cancel=True,
        )
        if not confirmed:
            self._set_status(u"Delete cancelled.", "neutral")
            return

        deleted = []
        errors = []
        for it in items:
            try:
                with revit.Transaction(u"Delete Family Parameter"):
                    self.fm.RemoveParameter(it.Param)
                deleted.append(it.Name)
            except Exception as ex:
                errors.append(u"{} ({})".format(it.Name, ex))

        parts = []
        if deleted:
            parts.append(u"Deleted {} parameter(s).".format(len(deleted)))
        if errors:
            parts.append(u"Errors: {}.".format(u"  ".join(errors)))

        tone = "error" if errors else "ok"
        self._set_status(u"  ".join(parts) if parts else u"Nothing deleted.", tone if parts else "neutral")
        self._reload_parameter_items(select_name=None)
        self._populate_sort_groups()
        self._update_batch_panel()

    def _restore_column_layout(self):
        from pyrevit import script as _pyscript
        from System.Windows.Controls import DataGridLength
        cfg = _pyscript.get_config()
        try:
            col_layout = getattr(cfg, 'col_layout', None)
            if not col_layout:
                return
            saved = []
            for part in col_layout.split(u","):
                if u":" in part:
                    header, width = part.rsplit(u":", 1)
                    saved.append((header.strip(), float(width)))
            col_by_header = {str(c.Header): c for c in self.lstParameters.Columns}
            # Restore widths first
            for header, width in saved:
                col = col_by_header.get(header)
                if col is not None:
                    col.Width = DataGridLength(width)
            # Restore display order (assign left-to-right; WPF adjusts others)
            for new_idx, (header, _) in enumerate(saved):
                col = col_by_header.get(header)
                if col is not None and col.DisplayIndex != new_idx:
                    col.DisplayIndex = new_idx
        except Exception:
            pass

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

    # ── Import Parameters tab ─────────────────────────────────────────────

    def _ensure_import_lib(self):
        """Lazy-load family_param_transfer; store references as instance attrs."""
        if getattr(self, u"_import_lib_loaded", False):
            return
        import family_param_transfer as _fpt
        self._fpt_get_open         = _fpt.get_open_family_docs
        self._fpt_read_params      = _fpt.read_params_from_doc
        self._fpt_read_nested      = _fpt.read_nested_families_from_doc
        self._fpt_read_styles      = _fpt.read_object_styles_from_doc
        self._fpt_transfer_params  = _fpt.transfer_params_to_doc
        self._fpt_transfer_nested  = _fpt.transfer_nested_families_to_doc
        self._fpt_transfer_styles  = _fpt.transfer_object_styles_to_doc
        self._import_lib_loaded = True
        # full lists from source doc (raw, unfiltered)
        self._import_all_params = []
        self._import_all_nested = []
        self._import_all_styles = []

    def _populate_import_source_combo(self):
        """Populate cmbImportSource with currently open family docs."""
        self._ensure_import_lib()
        docs = self._fpt_get_open(app, exclude_title=doc.Title)
        current = self.cmbImportSource.SelectedItem
        current_title = current.Title if current is not None else None
        self.cmbImportSource.ItemsSource = docs
        self.cmbImportSource.DisplayMemberPath = u"Title"
        if docs:
            for d in docs:
                if d.Title == current_title:
                    self.cmbImportSource.SelectedItem = d
                    break
            else:
                self.cmbImportSource.SelectedIndex = 0
        else:
            self.cmbImportSource.SelectedIndex = -1
            self._import_all_params = []
            self._import_all_nested = []
            self._import_all_styles = []
            self._apply_import_params_filter()
            self._apply_import_nested_filter()
            self._apply_import_styles_filter()
            self.btnImportExecute.IsEnabled = False

    def _get_import_source_doc(self):
        """Return the currently selected source Document or None."""
        sel = self.cmbImportSource.SelectedItem
        if sel is None:
            return None
        try:
            if sel.IsValidObject and sel.IsFamilyDocument:
                return sel
        except Exception:
            pass
        return None

    def _load_import_source_data(self, source_doc):
        """Read params, nested families, and object styles from source_doc."""
        self._ensure_import_lib()
        try:
            col = self._fpt_read_params(source_doc)
            self._import_all_params = list(col)
        except Exception:
            self._import_all_params = []
        try:
            col = self._fpt_read_nested(source_doc)
            self._import_all_nested = list(col)
        except Exception:
            self._import_all_nested = []
        try:
            col = self._fpt_read_styles(source_doc, include_model=True,
                                        include_annotation=True, include_imports=False)
            self._import_all_styles = list(col)
        except Exception:
            self._import_all_styles = []
        parents = sorted({s.ParentCategory for s in self._import_all_styles})
        self.cmbImportStylesCategory.ItemsSource = [u"(All)"] + parents
        self.cmbImportStylesCategory.SelectedIndex = 0

    def _apply_import_params_filter(self):
        from System.Collections.ObjectModel import ObservableCollection
        search = (self.txtImportSearch.Text or u"").strip().lower()
        hide_nt = bool(self.chkImportHideNonTransferable.IsChecked)
        filtered = []
        for item in self._import_all_params:
            if hide_nt and item.IsLikelyNonTransferable:
                continue
            if search and search not in item.Name.lower():
                continue
            filtered.append(item)
        col = ObservableCollection[object]()
        for item in filtered:
            col.Add(item)
        self.lstImportParams.ItemsSource = col
        self.txtImportParamsCount.Text = u"{} parameters shown".format(len(filtered))

    def _apply_import_nested_filter(self):
        from System.Collections.ObjectModel import ObservableCollection
        col = ObservableCollection[object]()
        for item in self._import_all_nested:
            col.Add(item)
        self.lstImportNested.ItemsSource = col

    def _apply_import_styles_filter(self):
        from System.Collections.ObjectModel import ObservableCollection
        inc_model = bool(self.chkImportStylesModel.IsChecked)
        inc_anno  = bool(self.chkImportStylesAnnotation.IsChecked)
        inc_imp   = bool(self.chkImportStylesImported.IsChecked)
        sel_parent = self.cmbImportStylesCategory.SelectedItem
        parent_filter = None if (sel_parent is None or sel_parent == u"(All)") else sel_parent
        filtered = []
        for item in self._import_all_styles:
            if item.StyleType == u"Model"      and not inc_model:  continue
            if item.StyleType == u"Annotation" and not inc_anno:   continue
            if item.StyleType == u"Import"     and not inc_imp:    continue
            if parent_filter and item.ParentCategory != parent_filter: continue
            filtered.append(item)
        col = ObservableCollection[object]()
        for item in filtered:
            col.Add(item)
        self.lstImportStyles.ItemsSource = col

    def _update_import_execute_button(self):
        src = self._get_import_source_doc()
        self.btnImportExecute.IsEnabled = (src is not None)

    def on_import_source_changed(self, sender, args):
        if not getattr(self, u"_import_lib_loaded", False):
            return
        src = self._get_import_source_doc()
        if src is None:
            self._import_all_params = []
            self._import_all_nested = []
            self._import_all_styles = []
            self._apply_import_params_filter()
            self._apply_import_nested_filter()
            self._apply_import_styles_filter()
            self.btnImportExecute.IsEnabled = False
            return
        self._load_import_source_data(src)
        self._apply_import_params_filter()
        self._apply_import_nested_filter()
        self._apply_import_styles_filter()
        self._update_import_execute_button()

    def on_import_refresh(self, sender, args):
        self._populate_import_source_combo()

    def on_import_subtab_changed(self, sender, args):
        from System.Windows.Controls import TabItem
        if not args.AddedItems or not isinstance(args.AddedItems[0], TabItem):
            return
        args.Handled = True

    def on_import_search_changed(self, sender, args):
        self._apply_import_params_filter()

    def on_import_filter_changed(self, sender, args):
        self._apply_import_params_filter()
        self._apply_import_styles_filter()

    def on_import_style_category_changed(self, sender, args):
        self._apply_import_styles_filter()

    def on_import_params_select_all(self, sender, args):
        src = self.lstImportParams.ItemsSource
        if src:
            for item in src:
                item.IsSelected = True

    def on_import_params_select_none(self, sender, args):
        src = self.lstImportParams.ItemsSource
        if src:
            for item in src:
                item.IsSelected = False

    def on_import_nested_select_all(self, sender, args):
        src = self.lstImportNested.ItemsSource
        if src:
            for item in src:
                item.IsSelected = True

    def on_import_nested_select_none(self, sender, args):
        src = self.lstImportNested.ItemsSource
        if src:
            for item in src:
                item.IsSelected = False

    def on_import_styles_select_all(self, sender, args):
        src = self.lstImportStyles.ItemsSource
        if src:
            for item in src:
                item.IsSelected = True

    def on_import_styles_select_none(self, sender, args):
        src = self.lstImportStyles.ItemsSource
        if src:
            for item in src:
                item.IsSelected = False

    def on_import_execute(self, sender, args):
        self._ensure_import_lib()
        src = self._get_import_source_doc()
        if src is None:
            self._set_status(u"No source family selected.", u"error")
            return

        hide_nt = bool(self.chkImportHideNonTransferable.IsChecked)
        sel_params = [i for i in self._import_all_params
                      if i.IsSelected and not (hide_nt and i.IsLikelyNonTransferable)]
        sel_nested = [i for i in self._import_all_nested if i.IsSelected]
        sel_styles = [i for i in self._import_all_styles if i.IsSelected]

        if not sel_params and not sel_nested and not sel_styles:
            self._set_status(u"Nothing selected to import.", u"error")
            return

        tf   = bool(self.chkImportTransferFormulas.IsChecked)
        tv   = bool(self.chkImportTransferValues.IsChecked)
        ov   = bool(self.chkImportOverwriteValues.IsChecked)
        cm   = bool(self.chkImportCreateMaterials.IsChecked)
        mo   = bool(self.chkImportNestedMissingOnly.IsChecked)
        re_x = bool(self.chkImportNestedReloadExisting.IsChecked)

        from pyrevit import script as _pyscript
        _out = _pyscript.get_output()
        _out.print_md(u"# Import Parameters Results")
        _out.print_md(u"**Source:** {}  \n**Target:** {}".format(src.Title, doc.Title))
        _out.print_md(u"---")

        if sel_params:
            res = self._fpt_transfer_params(app, src, sel_params, doc, tf, tv, ov, cm)
            _out.print_md(u"## Parameters")
            for name, _ in res[u"transferred"]:
                _out.print_md(u"- \u2713 {}".format(name))
            for name, reason in res[u"skipped"]:
                _out.print_md(u"- \u25e6 {} \u2014 _{}_ ".format(name, reason))
            for name, reason in res[u"failed"]:
                _out.print_md(u"- \u2717 {} \u2014 _{}_ ".format(name, reason))
            if res[u"formula_applied"]:
                _out.print_md(u"**Formulas applied:** {}".format(len(res[u"formula_applied"])))
            for name, reason in res[u"formula_failed"]:
                _out.print_md(u"- \u2717 formula {} \u2014 _{}_ ".format(name, reason))
            if res[u"values_copied"]:
                _out.print_md(u"**Values copied:** {}".format(len(res[u"values_copied"])))
            for name, reason in res[u"values_failed"]:
                _out.print_md(u"- \u2717 value {} \u2014 _{}_ ".format(name, reason))
            if res[u"materials_created"]:
                _out.print_md(u"**Materials created:** {}".format(
                    u", ".join(res[u"materials_created"])))
            for note in res[u"preload_notes"]:
                _out.print_md(u"- \u25e6 {}".format(note))

        if sel_nested:
            nres = self._fpt_transfer_nested(src, sel_nested, doc, mo, re_x)
            _out.print_md(u"## Nested Families")
            for name, _ in nres[u"loaded"]:
                _out.print_md(u"- \u2713 {} (loaded)".format(name))
            for name, _ in nres[u"reloaded"]:
                _out.print_md(u"- \u2713 {} (reloaded)".format(name))
            for name, reason in nres[u"skipped"]:
                _out.print_md(u"- \u25e6 {} \u2014 _{}_ ".format(name, reason))
            for name, reason in nres[u"failed"]:
                _out.print_md(u"- \u2717 {} \u2014 _{}_ ".format(name, reason))

        if sel_styles:
            sres = self._fpt_transfer_styles(src, sel_styles, doc)
            _out.print_md(u"## Object Styles")
            for name, _ in sres[u"created"]:
                _out.print_md(u"- \u2713 {} (created)".format(name))
            for name, _ in sres[u"copied"]:
                _out.print_md(u"- \u2713 {} (graphics copied)".format(name))
            for name, reason in sres[u"skipped"]:
                _out.print_md(u"- \u25e6 {} \u2014 _{}_ ".format(name, reason))
            for name, reason in sres[u"failed"]:
                _out.print_md(u"- \u2717 {} \u2014 _{}_ ".format(name, reason))

        self._reload_parameter_items(select_name=None)
        self._set_status(u"Import complete. See output window for details.", u"ok")

    def on_close(self, sender, args):
        from pyrevit import script as _pyscript
        cfg = _pyscript.get_config()
        cfg.win_left = self.Left
        cfg.win_top = self.Top
        cfg.win_width = self.Width
        cfg.win_height = self.Height
        try:
            ordered = sorted(self.lstParameters.Columns, key=lambda c: c.DisplayIndex)
            cfg.col_layout = u",".join(
                u"{}:{:.0f}".format(c.Header, c.ActualWidth) for c in ordered
            )
        except Exception:
            pass
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
        g = _get_group_type(fp.Definition)
    except Exception:
        return ""

    if g is None:
        return ""
    return _label_for_group(g)


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


def _get_group_type(definition):
    """Return the parameter's group value as a ForgeTypeId (2023+) or BuiltInParameterGroup.

    Uses GetGroupTypeId() introduced in Revit 2023; falls back to the
    ParameterGroup property which was removed in Revit 2025.
    """
    try:
        return definition.GetGroupTypeId()
    except AttributeError:
        return definition.ParameterGroup


def _default_group_type_general():
    """Return the 'General' parameter group value compatible with the running Revit version."""
    try:
        from Autodesk.Revit.DB import GroupTypeId
        return GroupTypeId.General
    except Exception:
        pass
    if _BIPG_AVAILABLE:
        return BuiltInParameterGroup.PG_GENERAL
    return None


def _infer_subexpr_datatype(subexpr_text, all_items):
    """
    Infer the data type that a subexpression produces by looking at the data
    types of the parameters it references and applying simple dimensional rules.

    Handles common cases:
    - LengthParam / 2          => Length
    - LengthParam + LengthParam2 => Length
    - LengthParam * NumberParam  => Length
    - LengthParam / LengthParam2 => ambiguous; returns Length (user can override)
    Returns None if no parameter references are found.
    """
    import re as _re

    def _cat(label_lower):
        if 'length'  in label_lower: return 'length'
        if 'volume'  in label_lower: return 'volume'
        if 'area'    in label_lower: return 'area'
        if 'angle'   in label_lower: return 'angle'
        if 'number'  in label_lower: return 'number'
        if 'integer' in label_lower: return 'number'
        if 'count'   in label_lower: return 'number'
        return 'other'

    # Build list of (name, dtype_value, category) sorted longest-name-first
    # so longer param names are matched before shorter ones that might be substrings.
    candidates = []
    for it in all_items:
        if not it.Name:
            continue
        dtype = _get_data_type(it.Param.Definition)
        if dtype is None:
            continue
        label = it.DataTypeLabel.lower() if it.DataTypeLabel else str(dtype).lower()
        candidates.append((it.Name, dtype, _cat(label)))
    candidates.sort(key=lambda x: -len(x[0]))

    # Identify which parameters are actually referenced (word-boundary match)
    referenced = []
    for name, dtype, cat in candidates:
        pat = r'(?<![A-Za-z0-9_])' + _re.escape(name) + r'(?![A-Za-z0-9_])'
        if _re.search(pat, subexpr_text):
            referenced.append((name, dtype, cat))

    if not referenced:
        return None

    # Separate dimensional vs numeric/dimensionless params
    dimensional = [(n, dt, c) for n, dt, c in referenced if c not in ('number', 'other')]

    if not dimensional:
        return referenced[0][1]   # all numeric — return first param's type

    # If only one dimensional category appears, return that type
    cats_present = {c for _, _, c in dimensional}
    if len(cats_present) == 1:
        return dimensional[0][1]

    # Multiple different dimensional types — return the highest-dimensional one
    for dom_cat in ('volume', 'area', 'length', 'angle'):
        for _, dt, c in dimensional:
            if c == dom_cat:
                return dt

    return referenced[0][1]


def _camel_to_words(s):
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", s)
    s = re.sub(r"(?<=[A-Z]{2})(?=[A-Z][a-z])", " ", s)
    return s.strip()


_STRUCTURAL_PT_NAMES = frozenset((
    "LoadClassification", "Mass", "Force", "LinearForce", "AreaForce",
    "Moment", "LinearMoment", "Stress", "UnitWeight", "Weight",
    "WeightPerUnitLength", "MomentOfInertia", "WarpingConstant",
    "SectionModulus", "SectionArea", "SectionDimension",
    "ReinforcementCover", "ReinforcementArea",
    "ReinforcementAreaperUnitLength", "ReinforcementAreaPerUnitLength",  # both capitalisations
    "ReinforcementSpacing", "ReinforcementVolume", "BarDiameter",
    "CrackWidth", "DisplacementDeflection", "Energy",
    "StructuralFrequency", "Period", "Pulsation", "Acceleration",
    "LinearMass", "LinearMassMomentOfInertia",
    "AreaSpringCoefficient", "LineSpringCoefficient", "PointSpringCoefficient",
    "RotationalLineSpringCoefficient", "RotationalPointSpringCoefficient",
    # Items that fall through the prefix checks but belong to Structural
    "AreaForcePerLength", "ForcePerLength", "ForceLengthPerAngle",
    "LinearForcePerLength", "LinearForceLengthPerAngle",
    "MassPerUnitArea", "MassPerUnitLength",
    "ReinforcementLength", "Rotation", "SectionProperty",
    "StructuralVelocity", "SurfaceArea", "ThermalExpansion",
))

_ELECTRICAL_PT_NAMES = frozenset((
    "ColorTemperature", "NumberOfPoles", "WireSize",
))

_PIPING_PT_NAMES = frozenset((
    "FixtureUnit", "PipeDimension", "PipeInsulationThickness",
    "PipeMass", "PipeMassPerUnitLength", "PipeSize",
))

_ENERGY_PT_NAMES = frozenset((
    "ThermalResistance", "ThermalMass", "ThermalConductivity",
    "SpecificHeat", "SpecificHeatOfVaporization", "Permeability",
    "IsothermalMoistureCapacity", "DiffusionCoefficient",
    "HeatTransferCoefficient", "AirflowDensity",
    "ThermalGradientCoefficientForMoistureCapacity", "MoistureDiffusivity",
))

_PT_LABEL_OVERRIDES = {
    "Boolean": "Yes/No",
    "YesNo": "Yes/No",
    "FamilyType": "<Family Type...>",
    "MultilineText": "Multiline Text",
    "MassDensity": "Mass Density",
    "URL": "URL",
    "HVACAirflowDividedByVolume": "Airflow / Volume",
    "HVACCoolingLoadDividedByArea": "Cooling Load / Area",
    "HVACCoolingLoadDividedByVolume": "Cooling Load / Volume",
    "HVACHeatingLoadDividedByArea": "Heating Load / Area",
    "HVACHeatingLoadDividedByVolume": "Heating Load / Volume",
    "DisplacementDeflection": "Displacement/Deflection",
    "ReinforcementAreaperUnitLength": "Reinforcement Area per Unit Length",
    "ReinforcementAreaPerUnitLength": "Reinforcement Area per Unit Length",
    "StructuralFrequency": "Frequency",
    "TimeInterval": "Time",
}


def _pt_discipline_and_label(name):
    discipline = "Common"
    raw = name
    for prefix, disc in (("HVAC", "HVAC"), ("Electrical", "Electrical"), ("Piping", "Piping")):
        if name.startswith(prefix) and len(name) > len(prefix):
            discipline = disc
            raw = name[len(prefix):]
            break
    else:
        if name in _STRUCTURAL_PT_NAMES:
            discipline = "Structural"
        elif name in _ENERGY_PT_NAMES:
            discipline = "Energy"
        elif name in _ELECTRICAL_PT_NAMES:
            discipline = "Electrical"
        elif name in _PIPING_PT_NAMES:
            discipline = "Piping"
    label = _PT_LABEL_OVERRIDES.get(name) or _camel_to_words(raw)
    return discipline, label


def _build_all_type_options_by_discipline():
    result = {}
    if _PT_AVAILABLE:
        # Pre-2025: enumerate the ParameterType enum
        try:
            for pt in System.Enum.GetValues(ParameterType):
                name = str(pt)
                if "invalid" in name.lower():
                    continue
                discipline, label = _pt_discipline_and_label(name)
                if not label:
                    continue
                if discipline not in result:
                    result[discipline] = []
                result[discipline].append(OptionItem(label, pt))
        except Exception:
            pass
    else:
        # 2025+: use Python attribute inspection on SpecTypeId.
        # SpecTypeId has nested classes (String, Int, Boolean, etc.) so we walk
        # them recursively with dir()/getattr() — more reliable in IronPython than
        # CLR reflection (GetNestedTypes/GetProperties/GetValue).
        try:
            from Autodesk.Revit.DB import SpecTypeId
            get_label_for_spec = getattr(LabelUtils, "GetLabelForSpec", None)
            if callable(get_label_for_spec):
                seen_specs = set()

                def _walk_spec(obj, depth=0):
                    if depth > 4:
                        return
                    for _aname in dir(obj):
                        if _aname.startswith('_'):
                            continue
                        try:
                            v = getattr(obj, _aname)
                        except Exception:
                            continue
                        if v is None:
                            continue
                        tid = getattr(v, 'TypeId', None)
                        if tid and isinstance(tid, str):
                            # It's a ForgeTypeId instance
                            if tid in seen_specs:
                                continue
                            seen_specs.add(tid)
                            try:
                                label = get_label_for_spec(v)
                                if not label:
                                    continue
                                tid_low = tid.lower()
                                if 'hvac' in tid_low:
                                    disc = 'HVAC'
                                elif 'electrical' in tid_low:
                                    disc = 'Electrical'
                                elif 'piping' in tid_low or 'plumbing' in tid_low:
                                    disc = 'Piping'
                                elif 'structural' in tid_low:
                                    disc = 'Structural'
                                elif 'energy' in tid_low:
                                    disc = 'Energy'
                                else:
                                    disc = 'Common'
                                result.setdefault(disc, []).append(OptionItem(label, v))
                            except Exception:
                                pass
                        elif isinstance(v, type) and depth < 3:
                            # Nested class — recurse into it
                            _walk_spec(v, depth + 1)

                _walk_spec(SpecTypeId)
        except Exception:
            pass
    for d in result:
        result[d].sort(key=lambda o: o.Label.lower())
    return result


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

    if _PT_AVAILABLE:
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
    else:
        # 2025+: use SpecTypeId equivalents
        try:
            from Autodesk.Revit.DB import SpecTypeId
            fallback = [
                (SpecTypeId.String.Text,    "Text"),
                (SpecTypeId.Number,         "Number"),
                (SpecTypeId.Length,         "Length"),
                (SpecTypeId.Boolean.YesNo,  "Yes/No"),
                (SpecTypeId.Int.Integer,    "Integer"),
            ]
            for spec, label in fallback:
                options.append(OptionItem(label, spec))
        except Exception:
            pass

    return options


# Groups shown by Revit in the Family Types / Edit Family Parameters dialog.
# Used to filter the full BuiltInParameterGroup enum down to the same subset.
_FAMILY_PARAM_GROUP_LABELS = frozenset((
    "analysis results",
    "analytical alignment",
    "analytical model",
    "constraints",
    "construction",
    "data",
    "dimensions",
    "division geometry",
    "electrical",
    "electrical - circuiting",
    "electrical - lighting",
    "electrical - loads",
    "electrical engineering",
    "energy analysis",
    "fire protection",
    "forces",
    "general",
    "graphics",
    "green building properties",
    "identity data",
    "ifc parameters",
    "layers",
    "materials and finishes",
    "mechanical",
    "mechanical - flow",
    "mechanical - loads",
    "model properties",
    "moments",
    "other",
    "overall legend",
    "phasing",
    "photometrics",
    "plumbing",
    "primary end",
    "rebar set",
    "releases / member forces",
    "secondary end",
    "segments and fittings",
    "set",
    "slab shape edit",
    "structural",
    "structural analysis",
    "text",
    "title text",
    "visibility",
))


def _build_group_options(fm):
    # Enumerate available parameter groups and keep only those Revit shows in the
    # Family Types / Create Parameter dialog.
    # Pre-2025: enumerate BuiltInParameterGroup enum.
    # 2025+:    enumerate GroupTypeId static ForgeTypeId properties via reflection.
    options = []
    seen = set()

    if _BIPG_AVAILABLE:
        try:
            for group in System.Enum.GetValues(BuiltInParameterGroup):
                key = str(group)
                if key in seen:
                    continue
                try:
                    label = LabelUtils.GetLabelFor(group)
                except Exception:
                    continue
                if not label or label == key:
                    continue
                if label.lower() not in _FAMILY_PARAM_GROUP_LABELS:
                    continue
                seen.add(key)
                options.append(OptionItem(label, group))
        except Exception:
            pass
    else:
        # Revit 2025+: use Python attribute inspection on GroupTypeId.
        # CLR reflection (GetProperties/GetValue) is unreliable in IronPython;
        # dir()/getattr() handles .NET static properties natively.
        # No label-set filter needed: every GroupTypeId property is a valid family group.
        try:
            from Autodesk.Revit.DB import GroupTypeId
            for _attr in dir(GroupTypeId):
                if _attr.startswith('_'):
                    continue
                try:
                    val = getattr(GroupTypeId, _attr)
                except Exception:
                    continue
                if val is None:
                    continue
                # ForgeTypeId instances expose a TypeId string; skip nested types/methods
                type_id_str = getattr(val, 'TypeId', None)
                if not type_id_str or not isinstance(type_id_str, str):
                    continue
                if type_id_str in seen:
                    continue
                label = _label_for_group(val)
                if not label:
                    continue
                seen.add(type_id_str)
                options.append(OptionItem(label, val))
        except Exception:
            pass

    options.sort(key=lambda o: o.Label.lower())
    return options


def _label_for_group(group):
    """Return a human-readable label for a group value (ForgeTypeId or BuiltInParameterGroup)."""
    if group is None:
        return ""
    # Revit 2025+: ForgeTypeId groups require GetLabelForGroup
    try:
        label = LabelUtils.GetLabelForGroup(group)
        if label:
            return label
    except Exception:
        pass
    # Pre-2025: BuiltInParameterGroup
    try:
        label = LabelUtils.GetLabelFor(group)
        if label:
            return label
    except Exception:
        pass
    return ""


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
    group_value = _get_group_type(source_param.Definition)
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


def _build_type_to_instance_chain(fm, start_name, max_depth=5):
    """
    Walk the formula-reference graph from *start_name* up to *max_depth* levels,
    collecting Type parameters whose formulas transitively reference it.

    Returns (chain, children_of):
      chain[0] = [start_name]
      chain[1] = Type params whose formula directly references start_name
      chain[2] = Type params whose formula references anything in chain[1]
      ...
      children_of: dict mapping each name -> list of names that directly reference it

    If no Type params reference the target, chain has only one element.
    """
    all_fps = list(_get_family_parameters(fm))
    chain = [[start_name]]
    seen = {start_name}
    children_of = {start_name: []}

    for _ in range(max_depth):
        current_names = chain[-1]
        next_level = []
        for fp in all_fps:
            name = _param_name(fp)
            if name in seen:
                continue
            if bool(getattr(fp, "IsInstance", False)):
                continue  # already instance — no conflict
            formula = _safe_formula(fp)
            if not formula:
                continue
            for ref_name in current_names:
                if _formula_references_parameter(formula, ref_name):
                    next_level.append(name)
                    seen.add(name)
                    children_of.setdefault(ref_name, []).append(name)
                    children_of.setdefault(name, [])
                    break
        if not next_level:
            break
        chain.append(next_level)

    return chain, children_of


def _build_instance_to_type_chain(fm, start_name, max_depth=5):
    """
    Walk the formula-reference graph DOWNWARD from *start_name* up to *max_depth* levels,
    collecting Instance parameters that the target's formula (transitively) references.
    These must be converted to Type before *start_name* can become Type.

    Returns (chain, children_of):
      chain[0] = [start_name]
      chain[1] = Instance params directly referenced in start_name's formula
      chain[2] = Instance params referenced in chain[1] params' formulas
      ...
      children_of: dict mapping each name -> list of Instance names it directly references

    If start_name's formula references no Instance params, chain has only one element.
    """
    all_fps = list(_get_family_parameters(fm))
    fp_by_name = {_param_name(fp): fp for fp in all_fps}

    chain = [[start_name]]
    seen = {start_name}
    children_of = {start_name: []}

    for _ in range(max_depth):
        current_names = chain[-1]
        next_level = []
        for parent_name in current_names:
            parent_fp = fp_by_name.get(parent_name)
            if parent_fp is None:
                continue
            formula = _safe_formula(parent_fp)
            if not formula:
                continue
            for fp in all_fps:
                name = _param_name(fp)
                if name in seen:
                    continue
                if not bool(getattr(fp, "IsInstance", False)):
                    continue  # only Instance params block a Type conversion
                if _formula_references_parameter(formula, name):
                    next_level.append(name)
                    seen.add(name)
                    children_of.setdefault(parent_name, []).append(name)
                    children_of.setdefault(name, [])
        if not next_level:
            break
        chain.append(next_level)

    return chain, children_of


def _format_chain_tree(start_name, children_of):
    """Render the dependency tree as a recursive box-draw hierarchy.

    Example (Param1 -> [Param2, Param5], Param2 -> [Param3], Param3 -> [Param4], Param5 -> [Param6]):
        'Param1'
        |- 'Param2'
        |  |- 'Param3'
        |  |  |- 'Param4'
        |- 'Param5'
           |- 'Param6'
    """
    lines = []
    lines.append(u"'{}'".format(start_name))

    def _render_children(parent, continuation):
        children = children_of.get(parent, [])
        for i, child in enumerate(children):
            is_last = (i == len(children) - 1)
            connector  = u"\u2514\u2500 " if is_last else u"\u251c\u2500 "
            child_cont = continuation + (u"   " if is_last else u"\u2502  ")
            lines.append(u"{}{}'{}'".format(continuation, connector, child))
            _render_children(child, child_cont)

    _render_children(start_name, u"")
    return u"\n".join(lines)


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
