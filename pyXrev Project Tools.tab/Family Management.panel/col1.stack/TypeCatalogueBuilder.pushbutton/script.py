# -*- coding: utf-8 -*-
"""
Type Catalogue Builder
----------------------
Builds a Revit type catalogue (.txt) from the Type parameters of the active
family.

Layout:
  Left listbox  â€” available Type params (no formula)
  Right listbox â€” selected params (ordered, become catalogue columns)
  DataGrid      â€” pre-populated from existing family types, fully editable
  Build button  â€” writes the .txt file

Header format:
  ,Param1##DATATYPE##UNITS,...

Validated against Revit 2021 (pre-2022 ParameterType/DUT_ path) and
Revit 2024 (ForgeTypeId path).
"""

import ast
import clr
import math
import operator
import os
import re

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Data")

from System import String
from System.Collections.ObjectModel import ObservableCollection
from System.Data import DataTable
from System.Windows import Style, DataTrigger, Setter
from System.Windows.Controls import DataGridCell, Control
from System.Windows.Data import Binding, IValueConverter
from System.Windows.Media import Color, SolidColorBrush

from Autodesk.Revit.DB import StorageType, UnitUtils
from pyrevit import forms, revit



# ---------------------------------------------------------------------------
# Revit context
# ---------------------------------------------------------------------------

doc = revit.doc
uiapp = __revit__  # noqa: F821

if not doc.IsFamilyDocument:
    forms.alert(
        "This tool must be run from inside the Family Editor.\n"
        "Open a family (.rfa) and try again.",
        exitscript=True,
    )

# ---------------------------------------------------------------------------
# Mapping tables â€” validated against Revit 2021 + 2024 diagnostic output
# ---------------------------------------------------------------------------

# ForgeTypeId substring â†’ catalogue DataType  (more-specific first)
_FORGE_DT_TOKENS = [
    (u"aec:length",   u"LENGTH"),
    (u"aec:angle",    u"ANGLE"),
    (u"aec:area",     u"AREA"),
    (u"aec:volume",   u"VOLUME"),
    (u"spec.bool",    u"YESNO"),    # autodesk.spec:spec.bool  (NOT ":bool")
    (u"spec.int",     u"INTEGER"),  # autodesk.spec:spec.int
    (u"spec.number",  u"NUMBER"),   # autodesk.spec:spec.number
    (u":number",      u"NUMBER"),   # autodesk.spec.measurable:number
]
# Everything else (string, url, material, familytype, image, currency) â†’ OTHER

# ForgeTypeId substring â†’ catalogue Units  (more-specific before less-specific
# so e.g. :squareFeet is matched before :feet)
_FORGE_UNIT_TOKENS = [
    (u":squareMillimeters",  u"SQUARE_MILLIMETERS"),
    (u":squareCentimeters",  u"SQUARE_CENTIMETERS"),
    (u":squareFeet",         u"SQUARE_FEET"),
    (u":squareInches",       u"SQUARE_INCHES"),
    (u":squareMeters",       u"SQUARE_METERS"),
    (u":cubicMillimeters",   u"CUBIC_MILLIMETERS"),
    (u":cubicCentimeters",   u"CUBIC_CENTIMETERS"),
    (u":cubicFeet",          u"CUBIC_FEET"),
    (u":cubicMeters",        u"CUBIC_METERS"),
    (u":millimeters",        u"MILLIMETERS"),
    (u":centimeters",        u"CENTIMETERS"),
    (u":meters",             u"METERS"),
    (u":feet",               u"FEET"),
    (u":fractionalInches",   u"FRACTIONAL_INCHES"),
    (u":decimalInches",      u"DECIMAL_INCHES"),
    (u":degrees",            u"DEGREES"),
    (u":radians",            u"RADIANS"),
    (u":liters",             u"LITERS"),
    (u":usGallons",          u"GALLONS_US"),
    (u":gallons",            u"GALLONS_US"),
    (u":kilonewtons",        u"KILONEWTONS"),
    (u":newtons",            u"NEWTONS"),
]

# pre-2022 ParameterType enum string â†’ catalogue DataType
_PARAM_TYPE_DT_MAP = {
    u"Length":  u"LENGTH",
    u"Angle":   u"ANGLE",
    u"Area":    u"AREA",
    u"Volume":  u"VOLUME",
    u"YesNo":   u"YESNO",
    u"Integer": u"INTEGER",
    u"Number":  u"NUMBER",
    # Text, URL, MultilineText, Material, FamilyType, Image, Currency â†’ OTHER
}

# pre-2022 DisplayUnitType string â†’ catalogue Units
_DUT_UNITS_MAP = {
    u"DUT_MILLIMETERS":            u"MILLIMETERS",
    u"DUT_CENTIMETERS":            u"CENTIMETERS",
    u"DUT_DECIMETERS":             u"METERS",
    u"DUT_METERS":                 u"METERS",
    u"DUT_METERS_CENTIMETERS":     u"METERS",
    u"DUT_DECIMAL_FEET":           u"FEET",
    u"DUT_FEET_FRACTIONAL_INCHES": u"FRACTIONAL_INCHES",
    u"DUT_FRACTIONAL_INCHES":      u"FRACTIONAL_INCHES",
    u"DUT_DECIMAL_INCHES":         u"DECIMAL_INCHES",
    u"DUT_DEGREES":                u"DEGREES",
    u"DUT_DEGREES_AND_MINUTES":    u"DEGREES",
    u"DUT_GRADS":                  u"DEGREES",
    u"DUT_RADIANS":                u"RADIANS",
    u"DUT_SQUARE_FEET":            u"SQUARE_FEET",
    u"DUT_SQUARE_INCHES":          u"SQUARE_INCHES",
    u"DUT_SQUARE_METERS":          u"SQUARE_METERS",
    u"DUT_SQUARE_MILLIMETERS":     u"SQUARE_MILLIMETERS",
    u"DUT_SQUARE_CENTIMETERS":     u"SQUARE_CENTIMETERS",
    u"DUT_CUBIC_FEET":             u"CUBIC_FEET",
    u"DUT_CUBIC_INCHES":           u"CUBIC_INCHES",
    u"DUT_CUBIC_METERS":           u"CUBIC_METERS",
    u"DUT_CUBIC_MILLIMETERS":      u"CUBIC_MILLIMETERS",
    u"DUT_CUBIC_CENTIMETERS":      u"CUBIC_CENTIMETERS",
    u"DUT_LITERS":                 u"LITERS",
    u"DUT_GALLONS_US":             u"GALLONS_US",
    u"DUT_KILONEWTONS":            u"KILONEWTONS",
    u"DUT_NEWTONS":                u"NEWTONS",
    # DUT_GENERAL, DUT_CURRENCY, DUT_PERCENTAGE â†’ "" (no units)
}

# DataTypes that carry display units
_DIMENSIONAL_DATATYPES = frozenset((u"LENGTH", u"ANGLE", u"AREA", u"VOLUME"))


# ---------------------------------------------------------------------------
# Catalogue info helper
# ---------------------------------------------------------------------------

def _catalogue_info(fp, doc):
    """
    Return (datatype_str, units_str, uid_for_conv).

    datatype_str  â€” catalogue header DataType token (LENGTH, YESNO, OTHER â€¦)
    units_str     â€” catalogue header Units token (MILLIMETERS, '', â€¦)
    uid_for_conv  â€” ForgeTypeId (2022+) or DisplayUnitType (pre-2022) needed
                    by UnitUtils.ConvertFromInternalUnits, or None.
    """
    # â”€â”€ 2022+ path: GetDataType() returns ForgeTypeId â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    try:
        dt = fp.Definition.GetDataType()
        if dt is not None:
            tid_lower = str(getattr(dt, u"TypeId", dt)).lower()

            datatype_str = u"OTHER"
            for tok, cat in _FORGE_DT_TOKENS:
                if tok in tid_lower:
                    datatype_str = cat
                    break

            units_str = u""
            uid_for_conv = None

            if datatype_str in _DIMENSIONAL_DATATYPES:
                try:
                    fmt = doc.GetUnits().GetFormatOptions(dt)
                    get_uid = getattr(fmt, u"GetUnitTypeId", None)
                    if callable(get_uid):
                        uid = get_uid()
                        if uid is not None:
                            uid_lower = str(getattr(uid, u"TypeId", uid)).lower()
                            for tok, u_str in _FORGE_UNIT_TOKENS:
                                if tok.lower() in uid_lower:
                                    units_str = u_str
                                    break
                            uid_for_conv = uid
                except Exception:
                    pass

            return datatype_str, units_str, uid_for_conv
    except AttributeError:
        pass  # pre-2022 â€” GetDataType() does not exist

    # â”€â”€ pre-2022 fallback: ParameterType enum + DisplayUnitType â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    try:
        pt_str = str(fp.Definition.ParameterType)
        datatype_str = _PARAM_TYPE_DT_MAP.get(pt_str, u"OTHER")

        units_str = u""
        uid_for_conv = None

        if datatype_str in _DIMENSIONAL_DATATYPES:
            try:
                unit_type = fp.Definition.UnitType
                fmt = doc.GetUnits().GetFormatOptions(unit_type)
                dut = fmt.DisplayUnits
                dut_str = str(dut)
                units_str = _DUT_UNITS_MAP.get(dut_str, u"")
                uid_for_conv = dut
            except Exception:
                pass

        return datatype_str, units_str, uid_for_conv
    except Exception:
        return u"OTHER", u"", None


# ---------------------------------------------------------------------------
# Parameter value reader
# ---------------------------------------------------------------------------

def _read_type_param_value(fm_type, fp, uid_for_conv):
    """Return the display-unit value string for fp under fm_type."""
    storage = fp.StorageType
    try:
        if storage == StorageType.Double:
            val = fm_type.AsDouble(fp)
            if uid_for_conv is not None:
                try:
                    display_val = UnitUtils.ConvertFromInternalUnits(val, uid_for_conv)
                    return _fmt_num(display_val)
                except Exception:
                    pass
            return _fmt_num(val)
        if storage == StorageType.Integer:
            return str(fm_type.AsInteger(fp))
        if storage == StorageType.String:
            s = fm_type.AsString(fp)
            return s if s is not None else u""
        if storage == StorageType.ElementId:
            return u""
    except Exception:
        return u""
    return u""


def _fmt_num(val):
    """Format a float for catalogue output â€” no trailing zeros, no sci notation."""
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
        return u"{:.4f}".format(f).rstrip(u"0").rstrip(u".")
    except Exception:
        return u""


def _csv_cell(s):
    """Quote a CSV cell only when it contains a delimiter, quote, or newline."""
    s = s or u""
    if u"," in s or u'"' in s or u"\n" in s:
        return u'"' + s.replace(u'"', u'""') + u'"'
    return s


# ---------------------------------------------------------------------------
# Formula engine
# ---------------------------------------------------------------------------
# Any cell value that starts with '=' is treated as a formula.
# Supported: arithmetic (+, -, *, /, **, %), column-name references,
#            and the functions abs(), round(), min(), max(), sqrt(), floor(), ceil().
# Column names are matched case-sensitively; names with special characters
# (spaces, hyphens …) are supported via a token-substitution step.

_FORMULA_OPS = {
    ast.Add:      operator.add,
    ast.Sub:      operator.sub,
    ast.Mult:     operator.mul,
    ast.Div:      operator.truediv,
    ast.Pow:      operator.pow,
    ast.Mod:      operator.mod,
    ast.FloorDiv: operator.floordiv,
    ast.USub:     operator.neg,
    ast.UAdd:     operator.pos,
}

_FORMULA_FUNCS = {
    u"abs":    abs,
    u"round":  round,
    u"min":    min,
    u"max":    max,
    u"sqrt":   math.sqrt,
    u"floor":  math.floor,
    u"ceil":   math.ceil,
    # concat(a, b, ...) joins all arguments as strings.
    # Numeric values are formatted the same way as catalogue output.
    u"concat": lambda *args: u"".join(
        _fmt_num(a) if isinstance(a, float) else (str(a) if a else u"")
        for a in args),
}


def _safe_id(name):
    """Turn an arbitrary column name into a valid Python identifier."""
    s = re.sub(r"[^A-Za-z0-9]", u"_", name)
    return (u"c_" + s) if (not s or s[0].isdigit()) else s


def _coerce_num(v):
    """Coerce v to float for arithmetic; raise ValueError if not possible."""
    if isinstance(v, float):
        return v
    try:
        return float(v)
    except (ValueError, TypeError):
        raise ValueError(u"'{}' is not a number".format(v))


def _val_to_str(v):
    """Format a formula value as a string (used by & concatenation)."""
    return _fmt_num(v) if isinstance(v, float) else (str(v) if v else u"")


def _eval_node(node, ctx):
    # Numeric literals
    if hasattr(ast, "Num") and isinstance(node, ast.Num):
        return float(node.n)
    if hasattr(ast, "Constant") and isinstance(node, ast.Constant):
        if isinstance(node.value, (int, float)):
            return float(node.value)
        if isinstance(node.value, str):
            return node.value           # string literal e.g. "prefix"
        raise ValueError(u"Unsupported literal type")
    # Python 2 string literal node
    if hasattr(ast, "Str") and isinstance(node, ast.Str):
        return node.s
    # Column name reference – returned as raw string; ops coerce as needed
    if isinstance(node, ast.Name):
        if node.id in ctx:
            return ctx[node.id]
        raise ValueError(u"Unknown column: {!r}".format(node.id))
    if isinstance(node, ast.BinOp):
        op_cls = type(node.op)
        left  = _eval_node(node.left,  ctx)
        right = _eval_node(node.right, ctx)
        # & is always string concatenation (like Excel/VBA)
        if op_cls == ast.BitAnd:
            return _val_to_str(left) + _val_to_str(right)
        if op_cls not in _FORMULA_OPS:
            raise ValueError(u"Unsupported operator")
        return _FORMULA_OPS[op_cls](_coerce_num(left), _coerce_num(right))
    if isinstance(node, ast.UnaryOp):
        op_cls = type(node.op)
        if op_cls not in _FORMULA_OPS:
            raise ValueError(u"Unsupported unary operator")
        return _FORMULA_OPS[op_cls](_coerce_num(_eval_node(node.operand, ctx)))
    if isinstance(node, ast.Call):
        if isinstance(node.func, ast.Name) and node.func.id in _FORMULA_FUNCS:
            args = [_eval_node(a, ctx) for a in node.args]
            result = _FORMULA_FUNCS[node.func.id](*args)
            # concat returns str; all other functions return numeric
            if node.func.id == u"concat":
                return result
            return float(result)
        raise ValueError(u"Function not allowed: {}".format(
            node.func.id if isinstance(node.func, ast.Name) else u"?"))
    raise ValueError(u"Unsupported expression: {}".format(type(node).__name__))


def evaluate_formula(formula_str, col_names, col_values):
    """
    Evaluate formula_str (the text after the leading '=').
    col_names  – list of parameter column names.
    col_values – matching list of current (already-resolved) string values.
    Returns float, or raises ValueError / SyntaxError.
    """
    # Build safe-id context and substitute column names in the expression.
    # Sort by length descending so longer names are replaced before shorter
    # ones that may be sub-strings of them.
    ctx = {}
    name_map = {}   # original name -> safe_id
    for name, val in sorted(zip(col_names, col_values), key=lambda x: -len(x[0])):
        sid = _safe_id(name)
        # Ensure uniqueness if two names map to the same safe_id
        while sid in ctx and name_map.get(name) != sid:
            sid += u"_"
        name_map[name] = sid
        ctx[sid] = val or u""  # raw string; arithmetic ops coerce to float lazily

    expr = formula_str
    for orig in sorted(name_map, key=len, reverse=True):
        expr = re.sub(re.escape(orig), name_map[orig], expr)

    tree = ast.parse(expr, mode=u"eval")
    return _eval_node(tree.body, ctx)


def _resolve_row_values(row, selected_params):
    """
    Return {col_name: value_string} for every param column in the row,
    with formulas evaluated.  Iterates up to 10 times so that formula
    columns can reference other formula columns.
    Non-evaluable formulas are left as-is (build step reports them).
    """
    col_names = [pi.Name for pi in selected_params]
    values = {pi.Name: str(row[pi.Name] or u"") for pi in selected_params}
    for _ in range(10):
        changed = False
        for pi in selected_params:
            val = values[pi.Name]
            if val.startswith(u"="):
                try:
                    raw = evaluate_formula(
                        val[1:], col_names,
                        [values[n] for n in col_names])
                    result = _fmt_num(raw) if isinstance(raw, float) else str(raw)
                    if result != val:
                        values[pi.Name] = result
                        changed = True
                except Exception:
                    pass  # leave formula string; build step reports it
        if not changed:
            break
    return values


# ---------------------------------------------------------------------------
# WPF IValueConverter – detects formula cells (value starts with '=')
# ---------------------------------------------------------------------------

class _FormulaCellConverter(IValueConverter):
    def Convert(self, value, targetType, parameter, culture):
        try:
            return value is not None and str(value).startswith(u"=")
        except Exception:
            return False

    def ConvertBack(self, value, targetType, parameter, culture):
        return value


# ---------------------------------------------------------------------------
# ParamInfo – lightweight wrapper around FamilyParameter
# ---------------------------------------------------------------------------

class ParamInfo(object):
    def __init__(self, fp, doc):
        self.FamilyParam = fp
        self.Name = fp.Definition.Name
        dt, units, uid = _catalogue_info(fp, doc)
        self.CatalogueDataType = dt
        self.CatalogueUnits = units
        self.UidForConv = uid

    @property
    def HeaderStr(self):
        if self.CatalogueUnits:
            return u"{}##{}##{}".format(
                self.Name, self.CatalogueDataType, self.CatalogueUnits)
        return u"{}##{}##".format(self.Name, self.CatalogueDataType)

    def __str__(self):
        return self.Name


# ---------------------------------------------------------------------------
# WPF Window
# ---------------------------------------------------------------------------

class TypeCatalogueBuilderWindow(forms.WPFWindow):

    def __init__(self, fm, doc):
        forms.WPFWindow.__init__(self, "TypeCatalogueBuilder.xaml")
        self._fm = fm
        self._doc = doc
        self._data_table = None
        self._formula_converter = _FormulaCellConverter()

        self._available_items = ObservableCollection[object]()
        self._selected_items = ObservableCollection[object]()

        # Wire formula events BEFORE first ItemsSource assignment so that
        # AutoGeneratingColumn fires with the handler already attached.
        self.dgTypes.AutoGeneratingColumn += self._on_auto_generating_column
        self.dgTypes.CellEditEnding += self._on_cell_edit_ending
        self.dgTypes.CurrentCellChanged += self._on_current_cell_changed

        self._populate_available()
        self.lstAvailable.ItemsSource = self._available_items
        self.lstSelected.ItemsSource = self._selected_items

        self._set_default_output_path()
        self._rebuild_datagrid()
        self._update_buttons()

    # â”€â”€ Setup helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _populate_available(self):
        """Fill lstAvailable with Type params (no formula), sorted by name."""
        self._available_items.Clear()
        selected_names = {pi.Name for pi in self._selected_items}
        params = []
        for fp in self._fm.Parameters:
            if fp.IsInstance:
                continue
            if fp.Formula:
                continue
            if fp.Definition.Name in selected_names:
                continue
            params.append(ParamInfo(fp, self._doc))
        params.sort(key=lambda p: p.Name.lower())
        for pi in params:
            self._available_items.Add(pi)

    def _set_default_output_path(self):
        family_name = (self._doc.Title or u"Family").replace(u".rfa", u"")
        folder = u""
        try:
            path = self._doc.PathName
            if path:
                folder = os.path.dirname(path)
        except Exception:
            pass
        default_path = (
            os.path.join(folder, family_name + u".txt") if folder
            else family_name + u".txt"
        )
        self.txtOutputPath.Text = default_path

    def _rebuild_datagrid(self):
        """
        Rebuild the DataTable from currently selected params.
        Preserves any cell values the user has already edited.
        """
        selected = list(self._selected_items)

        # Capture existing edits before rebuilding
        existing = {}   # {type_name: {col_name: value}}
        if self._data_table is not None:
            for row in self._data_table.Rows:
                tn = str(row[u"Type Name"] or u"")
                existing[tn] = {}
                for col in self._data_table.Columns:
                    cn = col.ColumnName
                    if cn != u"Type Name":
                        existing[tn][cn] = str(row[cn] or u"")

        dt = DataTable()
        dt.Columns.Add(u"Type Name", String)
        for pi in selected:
            dt.Columns.Add(pi.Name, String)

        for fm_type in self._fm.Types:
            row = dt.NewRow()
            tn = fm_type.Name or u""
            row[u"Type Name"] = tn
            for pi in selected:
                if tn in existing and pi.Name in existing[tn]:
                    row[pi.Name] = existing[tn][pi.Name]
                else:
                    row[pi.Name] = _read_type_param_value(
                        fm_type, pi.FamilyParam, pi.UidForConv)
            dt.Rows.Add(row)

        self._data_table = dt
        self.dgTypes.ItemsSource = dt.DefaultView

    def _update_buttons(self):
        has_avail  = self.lstAvailable.SelectedItems.Count > 0
        has_sel    = self.lstSelected.SelectedItems.Count > 0
        sel_count  = self._selected_items.Count
        any_avail  = self._available_items.Count > 0
        any_sel    = sel_count > 0

        sel_indices = sorted(
            self._selected_items.IndexOf(item)
            for item in self.lstSelected.SelectedItems
        ) if has_sel else []

        self.btnMoveAllRight.IsEnabled = any_avail
        self.btnMoveRight.IsEnabled    = has_avail
        self.btnMoveLeft.IsEnabled     = has_sel
        self.btnMoveAllLeft.IsEnabled  = any_sel
        self.btnMoveTop.IsEnabled      = has_sel and sel_indices[0] > 0
        self.btnMoveUp.IsEnabled       = has_sel and sel_indices[0] > 0
        self.btnMoveDown.IsEnabled     = has_sel and sel_indices[-1] < sel_count - 1
        self.btnMoveBottom.IsEnabled   = has_sel and sel_indices[-1] < sel_count - 1

    # -- Formula cell support ------------------------------------------------

    def _on_auto_generating_column(self, sender, args):
        """Apply a yellow cell highlight to every formula cell in the column."""
        col = args.Column
        col_name = str(col.Header)
        if col_name == u"Type Name":
            return
        cell_style = Style()
        cell_style.TargetType = DataGridCell
        trigger = DataTrigger()
        # Use WPF indexer-path syntax so column names with spaces work.
        trigger.Binding = Binding(u"[{}]".format(col_name))
        trigger.Binding.Converter = self._formula_converter
        trigger.Value = True
        trigger.Setters.Add(Setter(
            Control.BackgroundProperty,
            SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xCC))))
        cell_style.Triggers.Add(trigger)
        col.CellStyle = cell_style

    def _on_current_cell_changed(self, sender, args):
        """Refresh formula status label when the cursor moves."""
        self._update_formula_status()

    def _on_cell_edit_ending(self, sender, args):
        """Show live formula preview as the user commits a cell edit."""
        from System.Windows.Controls import TextBox as _WpfTextBox
        try:
            editing_elem = args.EditingElement
            if not isinstance(editing_elem, _WpfTextBox):
                return
            val = editing_elem.Text or u""
            if not val.startswith(u"="):
                self.lblFormulaStatus.Text = u""
                return
            col = args.Column
            prop = str(col.Header)
            item = args.Row.Item
            selected = list(self._selected_items)
            col_names = [pi.Name for pi in selected]
            col_values = [
                val if pi.Name == prop else str(item[pi.Name] or u"")
                for pi in selected
            ]
            try:
                result = evaluate_formula(val[1:], col_names, col_values)
                display = _fmt_num(result) if isinstance(result, float) else str(result)
                self.lblFormulaStatus.Text = u"\u2192  {}".format(display)
                self.lblFormulaStatus.Foreground = SolidColorBrush(
                    Color.FromRgb(0, 128, 0))
            except Exception as ex:
                self.lblFormulaStatus.Text = u"\u26a0  {}".format(str(ex))
                self.lblFormulaStatus.Foreground = SolidColorBrush(
                    Color.FromRgb(180, 0, 0))
        except Exception:
            pass

    def _update_formula_status(self):
        """Evaluate the current cell's formula and update the status label."""
        item = self.dgTypes.CurrentItem
        col = self.dgTypes.CurrentColumn
        if item is None or col is None:
            self.lblFormulaStatus.Text = u""
            return
        try:
            prop = str(col.Header)
            val = str(item[prop] or u"")
        except Exception:
            self.lblFormulaStatus.Text = u""
            return
        if not val.startswith(u"="):
            self.lblFormulaStatus.Text = u""
            return
        selected = list(self._selected_items)
        col_names = [pi.Name for pi in selected]
        col_values = [str(item[pi.Name] or u"") for pi in selected]
        try:
            result = evaluate_formula(val[1:], col_names, col_values)
            display = _fmt_num(result) if isinstance(result, float) else str(result)
            self.lblFormulaStatus.Text = u"\u2192  {}".format(display)
            self.lblFormulaStatus.Foreground = SolidColorBrush(
                Color.FromRgb(0, 128, 0))
        except Exception as ex:
            self.lblFormulaStatus.Text = u"\u26a0  {}".format(str(ex))
            self.lblFormulaStatus.Foreground = SolidColorBrush(
                Color.FromRgb(180, 0, 0))

    def on_available_selection_changed(self, sender, args):
        self._update_buttons()

    def on_selected_changed(self, sender, args):
        self._update_buttons()

    def on_available_double_click(self, sender, args):
        self._move_to_selected()

    def on_selected_double_click(self, sender, args):
        self._move_to_available()

    def on_move_all_right(self, sender, args):
        for item in list(self._available_items):
            self._selected_items.Add(item)
        self._available_items.Clear()
        self._rebuild_datagrid()
        self._update_buttons()

    def on_move_right(self, sender, args):
        self._move_to_selected()

    def on_move_left(self, sender, args):
        self._move_to_available()

    def on_move_all_left(self, sender, args):
        for item in list(self._selected_items):
            self._available_items.Add(item)
        self._selected_items.Clear()
        # Re-sort available list alphabetically
        sorted_items = sorted(list(self._available_items), key=lambda p: p.Name.lower())
        self._available_items.Clear()
        for item in sorted_items:
            self._available_items.Add(item)
        self._rebuild_datagrid()
        self._update_buttons()

    def on_move_top(self, sender, args):
        indices = sorted(
            self._selected_items.IndexOf(item)
            for item in self.lstSelected.SelectedItems
        )
        if not indices or indices[0] == 0:
            return
        items = [self._selected_items[i] for i in indices]
        for item in items:
            self._selected_items.Remove(item)
        for pos, item in enumerate(items):
            self._selected_items.Insert(pos, item)
        for item in items:
            self.lstSelected.SelectedItems.Add(item)
        self._rebuild_datagrid()
        self._update_buttons()

    def on_move_up(self, sender, args):
        indices = sorted(
            self._selected_items.IndexOf(item)
            for item in self.lstSelected.SelectedItems
        )
        if not indices or indices[0] == 0:
            return
        for idx in indices:
            item = self._selected_items[idx]
            self._selected_items.RemoveAt(idx)
            self._selected_items.Insert(idx - 1, item)
        for item in [self._selected_items[i - 1] for i in indices]:
            self.lstSelected.SelectedItems.Add(item)
        self._rebuild_datagrid()
        self._update_buttons()

    def on_move_down(self, sender, args):
        indices = sorted(
            (self._selected_items.IndexOf(item)
             for item in self.lstSelected.SelectedItems),
            reverse=True
        )
        last = self._selected_items.Count - 1
        if not indices or indices[0] >= last:
            return
        for idx in indices:
            item = self._selected_items[idx]
            self._selected_items.RemoveAt(idx)
            self._selected_items.Insert(idx + 1, item)
        for item in [self._selected_items[i + 1] for i in reversed(indices)]:
            self.lstSelected.SelectedItems.Add(item)
        self._rebuild_datagrid()
        self._update_buttons()

    def on_move_bottom(self, sender, args):
        indices = sorted(
            self._selected_items.IndexOf(item)
            for item in self.lstSelected.SelectedItems
        )
        last = self._selected_items.Count - 1
        if not indices or indices[-1] >= last:
            return
        items = [self._selected_items[i] for i in indices]
        for item in items:
            self._selected_items.Remove(item)
        for item in items:
            self._selected_items.Add(item)
        for item in items:
            self.lstSelected.SelectedItems.Add(item)
        self._rebuild_datagrid()
        self._update_buttons()

    def _move_to_selected(self):
        items = list(self.lstAvailable.SelectedItems)
        if not items:
            return
        for item in items:
            self._available_items.Remove(item)
            self._selected_items.Add(item)
        self._rebuild_datagrid()
        self._update_buttons()

    def _move_to_available(self):
        items = list(self.lstSelected.SelectedItems)
        if not items:
            return
        for item in items:
            self._selected_items.Remove(item)
            # Re-insert into available list in alphabetical position
            name_lower = item.Name.lower()
            insert_at = self._available_items.Count
            for i in range(self._available_items.Count):
                if self._available_items[i].Name.lower() > name_lower:
                    insert_at = i
                    break
            self._available_items.Insert(insert_at, item)
        self._rebuild_datagrid()
        self._update_buttons()

    def on_browse(self, sender, args):
        from Microsoft.Win32 import SaveFileDialog
        dlg = SaveFileDialog()
        dlg.Title = u"Save Type Catalogue"
        dlg.Filter = u"Text files (*.txt)|*.txt|All files (*.*)|*.*"
        dlg.DefaultExt = u"txt"
        current = (self.txtOutputPath.Text or u"").strip()
        if current:
            try:
                dlg.InitialDirectory = os.path.dirname(current) or u""
                dlg.FileName = os.path.basename(current)
            except Exception:
                pass
        if dlg.ShowDialog() == True:  # noqa: E712  (Nullable<bool> comparison)
            self.txtOutputPath.Text = dlg.FileName

    def on_build(self, sender, args):
        output_path = (self.txtOutputPath.Text or u"").strip()
        if not output_path:
            forms.alert(u"Please specify an output file path.", exitscript=False)
            return

        selected = list(self._selected_items)
        if not selected:
            forms.alert(u"Select at least one parameter column.", exitscript=False)
            return

        if self._data_table is None or self._data_table.Rows.Count == 0:
            ok = forms.alert(
                u"No family types are defined.\n"
                u"The catalogue will contain only a header row.\nContinue?",
                ok=True, cancel=True,
            )
            if not ok:
                return

        # Header row - first column is blank per Revit type catalogue format
        header_parts = [u""]
        header_parts += [pi.HeaderStr for pi in selected]
        lines = [u",".join(header_parts)]

        # Data rows — evaluate formulas before writing
        formula_errors = []
        if self._data_table is not None:
            for row in self._data_table.Rows:
                type_name = str(row[u"Type Name"] or u"")
                resolved = _resolve_row_values(row, selected)
                parts = [_csv_cell(type_name)]
                for pi in selected:
                    val = resolved[pi.Name]
                    if val.startswith(u"="):
                        formula_errors.append(
                            u"  {} / {}: {}".format(type_name, pi.Name, val))
                    parts.append(_csv_cell(val))
                lines.append(u",".join(parts))

        if formula_errors:
            msg = (u"The following formulas could not be evaluated:\n"
                   + u"\n".join(formula_errors[:10])
                   + u"\n\nWrite formula text as-is and proceed?")
            if not forms.alert(msg, ok=True, cancel=True):
                return

        try:
            with open(output_path, u"w") as f:
                f.write(u"\n".join(lines))
        except Exception as ex:
            forms.alert(u"Error saving file:\n{}".format(ex), exitscript=False)
            return

        forms.alert(u"Type catalogue saved:\n{}".format(output_path), title=u"Saved")
        self.Close()

    def on_close(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------------
# Fill Down helper
# ---------------------------------------------------------------------------
# The DataGrid is bound to DataTable.DefaultView, so each item is a
# DataRowView.  Values must be read/written with row[columnName].
# With SelectionUnit="Cell", selected rows come from grid.SelectedCells.

def _fill_down(grid):
    current = grid.CurrentCell
    col = current.Column
    if col is None:
        return
    prop = str(col.Header)

    # Collect rows that have a selected cell in this column, preserving grid order
    sel_rows = [c.Item for c in grid.SelectedCells if c.Column is col]
    if len(sel_rows) < 2:
        return  # nothing to fill into

    # First matching row in grid order is the source; fill the rest
    source_value = None
    found_source = False
    for row in grid.Items:
        if not found_source:
            for s in sel_rows:
                if s is row:
                    source_value = row[prop]
                    found_source = True
                    break
        else:
            for s in sel_rows:
                if s is row:
                    row[prop] = source_value
                    break

    grid.Items.Refresh()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

win = TypeCatalogueBuilderWindow(doc.FamilyManager, doc)

win.miFillDown.Click += lambda s, e: _fill_down(win.dgTypes)

win.ShowDialog()