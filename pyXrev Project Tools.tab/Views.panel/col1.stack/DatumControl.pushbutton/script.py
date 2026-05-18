# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Grid, Level, View,
    DatumExtentType, DatumEnds,
    Line, XYZ,
    ViewType,
)
from pyrevit import revit, forms, script as _pyscript
import sys

_ENV_KEYS = {
    "top":    "DatumExtentsFromCrop_Top",
    "bottom": "DatumExtentsFromCrop_Bottom",
    "left":   "DatumExtentsFromCrop_Left",
    "right":  "DatumExtentsFromCrop_Right",
}
_BUBBLE_KEYS = {
    "top":    "DatumExtentsFromCrop_BubTop",
    "bottom": "DatumExtentsFromCrop_BubBottom",
    "left":   "DatumExtentsFromCrop_BubLeft",
    "right":  "DatumExtentsFromCrop_BubRight",
}

doc   = revit.doc
uidoc = revit.uidoc

ALLOWED_TYPES = (
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.Elevation,
    ViewType.Section,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
    ViewType.Detail,
)

# --------------------------------------------------------------------------
# Determine target views
# Prefer views selected in the Project Browser; fall back to active view.
# --------------------------------------------------------------------------
def _is_valid_view(v):
    return (
        isinstance(v, View)
        and not v.IsTemplate
        and v.ViewType in ALLOWED_TYPES
        and v.CropBoxActive
    )

selected_ids = list(uidoc.Selection.GetElementIds())
selected_views = [
    doc.GetElement(eid)
    for eid in selected_ids
    if _is_valid_view(doc.GetElement(eid))
]

if selected_views:
    target_views = selected_views
else:
    active = uidoc.ActiveGraphicalView
    if active.ViewType not in ALLOWED_TYPES:
        forms.alert(
            "This tool only works in plan, section, elevation, or detail views.",
            exitscript=True,
        )
    if not active.CropBoxActive:
        forms.alert(
            "The active view has no active crop box.\n"
            "Enable the crop box and try again.",
            exitscript=True,
        )
    target_views = [active]

# --------------------------------------------------------------------------
# WPF dialog
# --------------------------------------------------------------------------
class DatumExtentsForm(forms.WPFWindow):
    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self.btnOK.Click     += self.on_ok
        self.btnCancel.Click += self.on_cancel
        self.tglEditOffset.Checked   += self._on_toggle_offset
        self.tglEditOffset.Unchecked += self._on_toggle_offset
        self.tglEditBubbles.Checked   += self._on_toggle_bubbles
        self.tglEditBubbles.Unchecked += self._on_toggle_bubbles
        # Apply initial state
        self._on_toggle_offset(None, None)
        self._on_toggle_bubbles(None, None)

    def _on_toggle_offset(self, sender, args):
        enabled = bool(self.tglEditOffset.IsChecked)
        for name in ("txtOffsetTop", "txtOffsetBottom", "txtOffsetLeft", "txtOffsetRight"):
            getattr(self, name).IsEnabled = enabled

    def _on_toggle_bubbles(self, sender, args):
        enabled = bool(self.tglEditBubbles.IsChecked)
        for name in ("chkBubbleTop", "chkBubbleBottom", "chkBubbleLeft", "chkBubbleRight"):
            getattr(self, name).IsEnabled = enabled

    def on_ok(self, sender, args):
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


dlg = DatumExtentsForm("DatumControl.xaml")
dlg.txtOffsetTop.Text    = _pyscript.get_envvar(_ENV_KEYS["top"])    or "10"
dlg.txtOffsetBottom.Text = _pyscript.get_envvar(_ENV_KEYS["bottom"]) or "10"
dlg.txtOffsetLeft.Text   = _pyscript.get_envvar(_ENV_KEYS["left"])   or "10"
dlg.txtOffsetRight.Text  = _pyscript.get_envvar(_ENV_KEYS["right"])  or "10"
for side, key in _BUBBLE_KEYS.items():
    saved = _pyscript.get_envvar(key)
    if saved is not None:
        getattr(dlg, "chkBubble" + side.capitalize()).IsChecked = (saved == "1")
if not dlg.ShowDialog():
    sys.exit(0)

adjust_grids  = dlg.chkGrids.IsChecked
adjust_levels = dlg.chkLevels.IsChecked
edit_offset   = bool(dlg.tglEditOffset.IsChecked)
edit_bubbles  = bool(dlg.tglEditBubbles.IsChecked)

if not adjust_grids and not adjust_levels:
    forms.alert("Both Grids and Levels are unchecked \u2014 nothing to do.", exitscript=True)

if not edit_offset and not edit_bubbles:
    forms.alert("Both Edit Offset and Edit Bubbles are off \u2014 nothing to do.", exitscript=True)

_offsets_mm = {}
for side, key in _ENV_KEYS.items():
    box = getattr(dlg, "txtOffset" + side.capitalize())
    if edit_offset:
        try:
            val = float(box.Text)
            if val < 0:
                raise ValueError()
        except ValueError:
            forms.alert(
                "'{}' offset must be a non-negative number.".format(side.capitalize()),
                exitscript=True,
            )
        _pyscript.set_envvar(key, box.Text)
    else:
        try:
            val = float(box.Text)
        except ValueError:
            val = 0.0
    _offsets_mm[side] = val

_bubbles = {}
for side, key in _BUBBLE_KEYS.items():
    checked = bool(getattr(dlg, "chkBubble" + side.capitalize()).IsChecked)
    _bubbles[side] = checked
    _pyscript.set_envvar(key, "1" if checked else "0")

# Warn if multiple views are about to be modified
if len(target_views) > 1:
    view_names = "\n".join("  • {}".format(v.Name) for v in target_views)
    if not forms.alert(
        "{} views are selected and will be updated:\n\n{}\n\nProceed?".format(
            len(target_views), view_names
        ),
        title="Multiple Views Selected",
        yes=True, no=True,
    ):
        sys.exit(0)

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def _crop_info(v):
    """Return (corners_world, local_x, local_y, bb) for the view's crop box.

    local_x  = view's right direction (world)
    local_y  = view's up direction (world)
    bb       = the BoundingBoxXYZ in view-local space
    corners  = four world-space corner points
    """
    bb = v.CropBox
    T  = bb.Transform
    corners = [
        T.OfPoint(XYZ(bb.Min.X, bb.Min.Y, 0.0)),
        T.OfPoint(XYZ(bb.Max.X, bb.Min.Y, 0.0)),
        T.OfPoint(XYZ(bb.Max.X, bb.Max.Y, 0.0)),
        T.OfPoint(XYZ(bb.Min.X, bb.Max.Y, 0.0)),
    ]
    local_x = T.BasisX   # right
    local_y = T.BasisY   # up
    return corners, local_x, local_y, bb


def _get_ref_curve(datum, v):
    for kind in (DatumExtentType.Model, DatumExtentType.ViewSpecific):
        try:
            curves = datum.GetCurvesInView(kind, v)
            if curves is not None and curves.Count > 0:
                return curves[0]
        except Exception:
            pass
    return None


def _set_datum_extent(datum, v, corners, local_x, local_y, bb, offsets_ft, bubbles,
                      edit_offset=True, edit_bubbles=True):
    """Set ViewSpecific extents and bubble visibility using per-side settings."""
    ref_curve = _get_ref_curve(datum, v)
    if ref_curve is None:
        return False

    direction = ref_curve.Direction
    ref_pt    = ref_curve.GetEndPoint(0)

    projs = [(c - ref_pt).DotProduct(direction) for c in corners]
    t_min_raw = min(projs)
    t_max_raw = max(projs)

    def _nearest_side(t):
        wp     = ref_pt + direction.Multiply(t)
        lx     = (wp - corners[0]).DotProduct(local_x)
        ly     = (wp - corners[0]).DotProduct(local_y)
        crop_w = bb.Max.X - bb.Min.X
        crop_h = bb.Max.Y - bb.Min.Y
        dists  = {
            "left":   lx,
            "right":  crop_w - lx,
            "bottom": ly,
            "top":    crop_h - ly,
        }
        return min(dists, key=dists.get)

    side_min = _nearest_side(t_min_raw)
    side_max = _nearest_side(t_max_raw)

    t_min = t_min_raw - offsets_ft[side_min]
    t_max = t_max_raw + offsets_ft[side_max]

    if (t_max - t_min) < 1e-6:
        t_min -= 1.0 / 304.8
        t_max += 1.0 / 304.8

    try:
        new_line = Line.CreateBound(
            ref_pt + direction.Multiply(t_min),
            ref_pt + direction.Multiply(t_max),
        )
    except Exception:
        return False

    datum.SetDatumExtentType(DatumEnds.End0, v, DatumExtentType.ViewSpecific)
    datum.SetDatumExtentType(DatumEnds.End1, v, DatumExtentType.ViewSpecific)

    if edit_offset:
        datum.SetCurveInView(DatumExtentType.ViewSpecific, v, new_line)

    if edit_bubbles:
        # End0 = t_min side, End1 = t_max side
        if bubbles[side_min]:
            datum.ShowBubbleInView(DatumEnds.End0, v)
        else:
            datum.HideBubbleInView(DatumEnds.End0, v)
        if bubbles[side_max]:
            datum.ShowBubbleInView(DatumEnds.End1, v)
        else:
            datum.HideBubbleInView(DatumEnds.End1, v)

    return True


# --------------------------------------------------------------------------
# Apply to all target views
# --------------------------------------------------------------------------
n_grids  = 0
n_levels = 0
failures = []

with revit.Transaction("Set Datum Extents from Crop"):
    for v in target_views:
        corners, local_x, local_y, bb = _crop_info(v)
        scale = v.Scale
        offsets_ft = {side: (mm * scale) / 304.8 for side, mm in _offsets_mm.items()}

        if adjust_grids:
            for g in FilteredElementCollector(doc, v.Id).OfClass(Grid).ToElements():
                try:
                    if _set_datum_extent(g, v, corners, local_x, local_y, bb, offsets_ft, _bubbles,
                                         edit_offset, edit_bubbles):
                        n_grids += 1
                    else:
                        failures.append("[{}] Grid '{}': could not read extent curve".format(v.Name, g.Name))
                except Exception as ex:
                    failures.append("[{}] Grid '{}': {}".format(v.Name, g.Name, ex))

        if adjust_levels:
            for lv in FilteredElementCollector(doc, v.Id).OfClass(Level).ToElements():
                try:
                    if _set_datum_extent(lv, v, corners, local_x, local_y, bb, offsets_ft, _bubbles,
                                          edit_offset, edit_bubbles):
                        n_levels += 1
                    else:
                        failures.append("[{}] Level '{}': could not read extent curve".format(v.Name, lv.Name))
                except Exception as ex:
                    failures.append("[{}] Level '{}': {}".format(v.Name, lv.Name, ex))


# --------------------------------------------------------------------------
# Summary
# --------------------------------------------------------------------------
parts = []
if n_grids:
    parts.append("{} grid{}".format(n_grids, "s" if n_grids != 1 else ""))
if n_levels:
    parts.append("{} level{}".format(n_levels, "s" if n_levels != 1 else ""))

view_label = (
    "across {} views".format(len(target_views))
    if len(target_views) > 1 else
    "in '{}'".format(target_views[0].Name)
)

msg = (
    "Set extents for {} {}.".format(" and ".join(parts), view_label)
    if parts else
    "No datums were updated."
)
if failures:
    msg += "\n\nFailed ({}):\n".format(len(failures)) + "\n".join(failures)

forms.alert(msg, title="Datum Extents from Crop")
