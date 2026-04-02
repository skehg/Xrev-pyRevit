# -*- coding: utf-8 -*-
"""
Purge Unused Parameters
-----------------------
Identifies family parameters that have no impact on the function of the family
and allows the user to selectively delete them.

Detection strategy (two-pass):
  Pass 1 – Dimension labels: any parameter assigned as a dimension label is
            directly in use.
  Pass 2 – Sub-transaction delete test: for every parameter not already flagged,
            attempt to remove it inside a sub-transaction that is immediately
            rolled back.  If Revit refuses (raises), the parameter is considered
            directly in use (drives visibility, an array count, etc.).

After the directly-used set is established a BFS walk follows the formula
dependency chain up to the user-specified depth.  Parameters that feed the
formula of a directly-used parameter are also unsafe to delete; parameters
that feed THOSE are unsafe at depth 2; and so on.

Any parameter not captured by this walk is shown in the results grid with
reverse-dependency columns D1…Dn.  Each depth column lists the other safe
parameters whose formulas reference this parameter at that hop distance,
giving the user visibility into knock-on effects within the safe pool.

Parameters are deleted in topological (leaf-first) order so that formula
dependencies between selected parameters do not cause Revit errors.
"""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from pyrevit import forms, revit
from family_param_utils import (
    MAX_DEPTH_CAP as _MAX_DEPTH_CAP,
    get_family_parameters as _get_family_parameters,
    safe_formula as _safe_formula,
    param_name as _param_name,
    group_label as _group_label,
    data_type_label as _data_type_label,
    formula_references_parameter as _formula_references_parameter,
    find_directly_used_params as _find_directly_used_params,
    build_depth_analysis as _build_depth_analysis,
    compute_reverse_deps as _compute_reverse_deps,
)
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Controls import DataGridTextColumn, DataGridLength, DataGridLengthUnitType
from System.Windows.Data import Binding as WpfBinding
from System.Windows.Media import SolidColorBrush, Color


doc = revit.doc
uiapp = __revit__   # noqa: F821
app = uiapp.Application


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class ParamResultItem(INotifyPropertyChanged):
    """One row in the results grid."""

    # Pre-declare all possible depth properties so WPF binding can find them
    # even before columns are dynamically added.
    D1  = ""
    D2  = ""
    D3  = ""
    D4  = ""
    D5  = ""
    D6  = ""
    D7  = ""
    D8  = ""
    D9  = ""
    D10 = ""

    def __init__(self, fp, rev_deps, max_depth):
        self._is_selected = True
        self._fp = fp
        self.Name = _param_name(fp)
        self.GroupLabel = _group_label(fp)
        self.InstanceTypeLabel = "Instance" if bool(getattr(fp, "IsInstance", False)) else "Type"
        self.DataTypeLabel = _data_type_label(fp)
        self.Formula = _safe_formula(fp)
        self.HasFormula = bool(self.Formula)

        # Populate depth columns
        for d in range(1, min(max_depth, _MAX_DEPTH_CAP) + 1):
            setattr(self, "D{}".format(d), rev_deps.get(d, ""))

        self._changed_handlers = []

    # INotifyPropertyChanged
    def add_PropertyChanged(self, handler):
        self._changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._changed_handlers:
            self._changed_handlers.remove(handler)

    def _notify(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for h in self._changed_handlers:
            h(self, args)

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify("IsSelected")

    @property
    def FamilyParameter(self):
        return self._fp


# ---------------------------------------------------------------------------
# WPF window
# ---------------------------------------------------------------------------

class PurgeParametersWindow(forms.WPFWindow):

    def __init__(self, family_manager):
        forms.WPFWindow.__init__(self, "PurgeParameters.xaml")
        self.fm = family_manager
        self._items = ObservableCollection[object]()
        self.dataGrid.ItemsSource = self._items
        self._current_depth = 5
        self._depth_columns_added = 0

        self._set_status("Click Analyse to scan the family for unused parameters.", "neutral")
        self._sync_depth_display()

        # Wire up selection-changed to enable/disable Delete button
        self._items.CollectionChanged += self._on_collection_changed

    # ------------------------------------------------------------------
    # Depth spinner helpers
    # ------------------------------------------------------------------

    def _sync_depth_display(self):
        self.txtDepth.Text = str(self._current_depth)

    def on_depth_down(self, sender, args):
        if self._current_depth > 1:
            self._current_depth -= 1
            self._sync_depth_display()

    def on_depth_up(self, sender, args):
        if self._current_depth < _MAX_DEPTH_CAP:
            self._current_depth += 1
            self._sync_depth_display()

    def _read_depth(self):
        try:
            v = int(self.txtDepth.Text.strip())
            return max(1, min(_MAX_DEPTH_CAP, v))
        except Exception:
            return self._current_depth

    # ------------------------------------------------------------------
    # Analysis
    # ------------------------------------------------------------------

    def on_analyse_clicked(self, sender, args):
        self._current_depth = self._read_depth()
        self._sync_depth_display()
        self._run_analysis(self._current_depth)

    def _run_analysis(self, depth):
        self._set_status("Analysing\u2026", "neutral")
        self.btnAnalyse.IsEnabled = False
        self.btnDelete.IsEnabled = False

        try:
            # Phase 1 – detect directly-used parameters
            self._set_status(
                "Testing each parameter (attempting delete + rollback per parameter) \u2014 "
                "this may take a moment for large families\u2026",
                "neutral",
            )
            directly_used = _find_directly_used_params(doc, self.fm)

            # Phase 2 – BFS depth analysis
            self._set_status("Tracing formula dependency chain (depth {0})\u2026".format(depth), "neutral")
            safe_params, _unsafe = _build_depth_analysis(self.fm, directly_used, depth)

            # Phase 3 – reverse dependencies within the safe pool
            rev_dep_map = _compute_reverse_deps(safe_params, depth)

            # Rebuild grid columns
            self._rebuild_depth_columns(depth)

            # Populate rows
            self._items.Clear()
            for fp in safe_params:
                name = _param_name(fp)
                per_param_deps = rev_dep_map.get(name, {})
                item = ParamResultItem(fp, per_param_deps, depth)
                item.PropertyChanged += self._on_item_property_changed
                self._items.Add(item)

            count = len(safe_params)
            all_count = len(_get_family_parameters(self.fm))
            self.txtParamCount.Text = (
                "{0} of {1} parameters are safe to delete".format(count, all_count)
                if count else "No unused parameters found at depth {0}".format(depth)
            )

            self._update_delete_button()

            if count:
                self._set_status(
                    "Analysis complete.  {0} parameter(s) identified as safe to delete.  "
                    "Tick the ones you want to remove and click Delete Selected.".format(count),
                    "ok",
                )
            else:
                self._set_status(
                    "No unused parameters found at depth {0}.  "
                    "Try reducing the depth to see parameters that are only indirectly used.".format(depth),
                    "neutral",
                )

        except Exception as ex:
            self._set_status("Error during analysis: {}".format(ex), "error")
        finally:
            self.btnAnalyse.IsEnabled = True

    def _rebuild_depth_columns(self, depth):
        """Remove any previous depth columns and add fresh ones for D1..Ddepth."""
        # Remove columns whose Header starts with "D" and is a depth column
        to_remove = []
        for col in self.dataGrid.Columns:
            header = getattr(col, "Header", None)
            if header is not None and str(header).startswith("D") and str(header)[1:].isdigit():
                to_remove.append(col)
        for col in to_remove:
            self.dataGrid.Columns.Remove(col)

        # Add new depth columns
        for d in range(1, depth + 1):
            col = DataGridTextColumn()
            col.Header = "D{}  (references D{} params)".format(d, d) if d > 1 else "D1  (direct references)"
            col.IsReadOnly = True
            col.Width = DataGridLength(180)
            binding = WpfBinding("D{}".format(d))
            col.Binding = binding
            self.dataGrid.Columns.Add(col)

        self._depth_columns_added = depth

    # ------------------------------------------------------------------
    # Selection helpers
    # ------------------------------------------------------------------

    def on_select_all(self, sender, args):
        for item in self._items:
            item.IsSelected = True
        self._update_delete_button()

    def on_deselect_all(self, sender, args):
        for item in self._items:
            item.IsSelected = False
        self._update_delete_button()

    def on_invert_selection(self, sender, args):
        for item in self._items:
            item.IsSelected = not item.IsSelected
        self._update_delete_button()

    def _on_collection_changed(self, sender, args):
        self._update_delete_button()

    def _on_item_property_changed(self, sender, args):
        if args.PropertyName == "IsSelected":
            self._update_delete_button()

    def _update_delete_button(self):
        selected_count = sum(1 for item in self._items if item.IsSelected)
        self.btnDelete.IsEnabled = selected_count > 0
        if selected_count:
            self.btnDelete.Content = "Delete Selected ({0})".format(selected_count)
        else:
            self.btnDelete.Content = "Delete Selected"

    # ------------------------------------------------------------------
    # Topological sort
    # ------------------------------------------------------------------

    def _topological_sort(self, selected_items):
        """Return *selected_items* in leaf-first deletion order.

        A parameter A must be deleted before parameter B if B's formula
        references A (so B is a "consumer" of A — delete A last relative to B).
        Equivalently: delete parameters whose names are NOT referenced by any
        remaining selected parameter first.
        """
        remaining = list(selected_items)
        ordered = []

        # Build name -> item map for the selection
        name_to_item = {item.Name: item for item in remaining}
        selected_names = set(name_to_item.keys())

        # Build dependency: deps[name] = set of selected names that THIS param
        # references in its own formula (i.e. must be deleted after those).
        deps = {}
        for item in remaining:
            own_formula = item.Formula
            deps[item.Name] = set()
            if own_formula:
                for other_name in selected_names:
                    if other_name != item.Name and _formula_references_parameter(own_formula, other_name):
                        deps[item.Name].add(other_name)

        # Iteratively extract nodes with no remaining dependencies
        max_iterations = len(remaining) + 1
        iteration = 0
        while remaining and iteration < max_iterations:
            iteration += 1
            leaves = [item for item in remaining if not deps.get(item.Name)]
            if not leaves:
                # Circular dependency among selected params — just append the rest
                ordered.extend(remaining)
                break
            for leaf in leaves:
                ordered.append(leaf)
                remaining.remove(leaf)
                # Remove this name from other items' dependency sets
                for other_item in remaining:
                    deps[other_item.Name].discard(leaf.Name)

        return ordered

    # ------------------------------------------------------------------
    # Deletion
    # ------------------------------------------------------------------

    def on_delete_clicked(self, sender, args):
        selected = [item for item in self._items if item.IsSelected]
        if not selected:
            return

        names_list = "\n".join("  - " + item.Name for item in selected[:20])
        if len(selected) > 20:
            names_list += "\n  \u2026 and {} more".format(len(selected) - 20)

        confirmed = forms.alert(
            "Delete the following {0} parameter(s)?\n\n{1}".format(len(selected), names_list),
            title="Confirm Deletion",
            ok=True,
            cancel=True,
        )
        if not confirmed:
            return

        ordered = self._topological_sort(selected)

        errors = []
        with revit.Transaction("Purge Unused Parameters"):
            for item in ordered:
                try:
                    self.fm.RemoveParameter(item.FamilyParameter)
                except Exception as ex:
                    errors.append("{}: {}".format(item.Name, ex))

        if errors:
            self._set_status(
                "Deleted {0} parameter(s).  {1} error(s): {2}".format(
                    len(ordered) - len(errors),
                    len(errors),
                    "; ".join(errors),
                ),
                "error",
            )
        else:
            self._set_status(
                "Successfully deleted {0} parameter(s).".format(len(ordered)),
                "ok",
            )

        # Re-run analysis with same depth to refresh the grid
        self._run_analysis(self._current_depth)

    # ------------------------------------------------------------------
    # Close
    # ------------------------------------------------------------------

    def on_close_clicked(self, sender, args):
        self.Close()

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def _set_status(self, message, tone):
        self.txtStatus.Text = message or ""
        colours = {
            "ok":      "#1A6B1A",
            "error":   "#B00020",
            "neutral": "#444444",
        }
        hex_col = colours.get(tone, "#444444").lstrip("#")
        r = int(hex_col[0:2], 16)
        g = int(hex_col[2:4], 16)
        b = int(hex_col[4:6], 16)
        self.txtStatus.Foreground = SolidColorBrush(Color.FromRgb(r, g, b))


# ---------------------------------------------------------------------------
# Guard + entry point
# ---------------------------------------------------------------------------

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

    dlg = PurgeParametersWindow(fm)
    dlg.ShowDialog()


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        forms.alert("Unexpected error:\n{}".format(ex), title="Purge Parameters Error")
