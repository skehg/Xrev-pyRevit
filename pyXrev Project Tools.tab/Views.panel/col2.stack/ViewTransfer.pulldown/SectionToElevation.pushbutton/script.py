# pylint: disable=E0401,W0613,C0103,C0111
# -*- coding: utf-8 -*-
"""Convert selected Section views to Elevation views.

Uses ElevationMarker.CreateElevationMarker (same mechanism as pyChilizer's
Room Elevations Plus) so the result is a genuine Elevation view. The marker
is placed at the section's origin and rotated to face the same direction.
Crop region, far clip, scale, phase, view template and view-owned detail
elements are all copied from the source section.

New elevations are prefixed with "Sec_". Original sections are kept.
"""
import sys
from Autodesk.Revit.DB import (
    BuiltInParameter,
    CopyPasteOptions,
    CurveLoop,
    DuplicateTypeAction,
    ElementId,
    ElementTransformUtils,
    ElevationMarker,
    FilteredElementCollector,
    IDuplicateTypeNamesHandler,
    Line,
    ViewFamilyType,
    ViewFamily,
    ViewSection,
    ViewType,
)
from pyrevit import revit, DB, forms
from pyrevit.framework import List

doc = revit.doc


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

class _CopyUseDestination(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


def _get_vft_by_family(document, view_family):
    for vft in FilteredElementCollector(document).OfClass(ViewFamilyType):
        try:
            if vft.ViewFamily == view_family:
                return vft
        except Exception:
            continue
    return None


def _get_floor_plan_view(document):
    """Return any floor plan view suitable for hosting an ElevationMarker."""
    for v in FilteredElementCollector(document).OfClass(DB.View).ToElements():
        if v.ViewType == ViewType.FloorPlan and not v.IsTemplate:
            return v
    return None


def _unique_name(base_name, existing_names):
    if base_name not in existing_names:
        return base_name
    counter = 2
    candidate = "{} ({})".format(base_name, counter)
    while candidate in existing_names:
        counter += 1
        candidate = "{} ({})".format(base_name, counter)
    return candidate


def _collect_view_owned_elements(document, view):
    # Whitelist of concrete types that are safe to copy across views.
    # Section markers, callout bubbles and elevation tags are system annotation
    # types that do NOT appear in this list, so they are never collected.
    _SAFE_TYPES = (
        DB.CurveElement,    # detail lines, arcs, splines (ViewSpecific)
        DB.FilledRegion,
        DB.TextNote,
        DB.Dimension,
        DB.IndependentTag,
        DB.FamilyInstance,  # detail components, generic annotations
        DB.RevisionCloud,
    )
    owned = []
    for el in FilteredElementCollector(document, view.Id).ToElements():
        try:
            if not (el.ViewSpecific and el.OwnerViewId == view.Id):
                continue
            if not isinstance(el, _SAFE_TYPES):
                continue
            owned.append(el.Id)
        except Exception:
            continue
    return owned


def _set_elevation_crop(section, new_elev):
    """Set crop region of new_elev to match the section's crop extents.

    Transforms the 4 crop corners from section local space to world space via
    CropBox.Transform, then applies them via SetCropShape. Falls back to
    directly writing CropBox Min/Max if SetCropShape fails.
    """
    cb = section.CropBox
    t  = cb.Transform
    min_x, max_x = cb.Min.X, cb.Max.X
    min_y, max_y = cb.Min.Y, cb.Max.Y

    bl = t.OfPoint(DB.XYZ(min_x, min_y, 0))
    tl = t.OfPoint(DB.XYZ(min_x, max_y, 0))
    tr = t.OfPoint(DB.XYZ(max_x, max_y, 0))
    br = t.OfPoint(DB.XYZ(max_x, min_y, 0))

    new_elev.CropBoxActive = True
    try:
        l1 = Line.CreateBound(bl, tl)
        l2 = Line.CreateBound(tl, tr)
        l3 = Line.CreateBound(tr, br)
        l4 = Line.CreateBound(br, bl)
        crop_loop = CurveLoop.Create(List[DB.Curve]([l1, l2, l3, l4]))
        new_elev.GetCropRegionShapeManager().SetCropShape(crop_loop)
    except Exception:
        try:
            cb_new = new_elev.CropBox
            cb_new.Min = DB.XYZ(min_x, min_y, cb_new.Min.Z)
            cb_new.Max = DB.XYZ(max_x, max_y, cb_new.Max.Z)
            new_elev.CropBox = cb_new
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Step 1 – Select sections
# ---------------------------------------------------------------------------

selection = revit.get_selection()
sections = [
    doc.GetElement(eid)
    for eid in selection.element_ids
    if isinstance(doc.GetElement(eid), ViewSection)
    and doc.GetElement(eid).ViewType == ViewType.Section
]

if not sections:
    sections = forms.select_views(
        title="Select Section Views to Convert",
        filterfunc=lambda v: (
            isinstance(v, ViewSection) and v.ViewType == ViewType.Section
        ),
        use_selection=False,
    )

if not sections:
    forms.alert("No Section views selected.", exitscript=True)

# ---------------------------------------------------------------------------
# Step 2 – Find Elevation ViewFamilyType
# ---------------------------------------------------------------------------

elev_vft = _get_vft_by_family(doc, ViewFamily.Elevation)
if not elev_vft:
    forms.alert(
        "No Elevation ViewFamilyType found in this project.",
        exitscript=True,
    )

# ---------------------------------------------------------------------------
# Step 3 – Find a floor plan view for ElevationMarker hosting
# ---------------------------------------------------------------------------

plan_view = _get_floor_plan_view(doc)
if not plan_view:
    forms.alert(
        "No Floor Plan view found. ElevationMarker requires a plan view to host the elevation.",
        exitscript=True,
    )

# ---------------------------------------------------------------------------
# Step 4 – Build set of existing view names (for uniqueness checks)
# ---------------------------------------------------------------------------

existing_names = set(
    v.Name
    for v in FilteredElementCollector(doc)
               .OfClass(DB.View)
               .WhereElementIsNotElementType()
               .ToElements()
    if not v.IsTemplate
)

# ---------------------------------------------------------------------------
# Step 5 – Convert
# ---------------------------------------------------------------------------

copy_options = CopyPasteOptions()
copy_options.SetDuplicateTypeNamesHandler(_CopyUseDestination())

created = []
skipped = []

with revit.Transaction("Convert Sections to Elevations"):
    for section in sections:
        sec_name = section.Name
        new_name = _unique_name("Sec_" + sec_name, existing_names)

        # --- Place ElevationMarker at the midpoint of the section cut line ---
        # section.Origin is the section head (one end); we want the horizontal
        # centre of the crop box on the cut plane so the marker lands at the
        # midpoint of the view extents (the blue-circle position in plan).
        _cb = section.CropBox
        _mid_x = (_cb.Min.X + _cb.Max.X) / 2.0
        marker_pos = _cb.Transform.OfPoint(DB.XYZ(_mid_x, 0.0, 0.0))
        try:
            new_marker = ElevationMarker.CreateElevationMarker(
                doc, elev_vft.Id, marker_pos, section.Scale
            )
            # Create all 4 slots so we can measure each actual ViewDirection
            # and keep only the one that best matches section.ViewDirection.
            _probes = {}
            for _i in range(4):
                try:
                    _probes[_i] = new_marker.CreateElevation(doc, plan_view.Id, _i)
                except Exception:
                    pass
            if not _probes:
                raise Exception("ElevationMarker.CreateElevation failed for all slots")
            doc.Regenerate()
            _sv = section.ViewDirection
            _best_idx = max(_probes, key=lambda i: _probes[i].ViewDirection.DotProduct(_sv))
            new_elev = _probes[_best_idx]
            # Delete the 3 unused slots via the ICollection overload which is
            # more reliable than individual doc.Delete calls inside a Python loop.
            _del_ids = List[ElementId]([v.Id for i, v in _probes.items() if i != _best_idx])
            if _del_ids:
                try:
                    doc.Delete(_del_ids)
                except Exception:
                    # Fallback: try one-by-one
                    for _i, _v in _probes.items():
                        if _i != _best_idx:
                            try:
                                doc.Delete(_v.Id)
                            except Exception:
                                pass
            doc.Regenerate()
        except Exception as ex:
            skipped.append((sec_name, str(ex)))
            continue

        # --- Name ---
        new_elev.Name = new_name
        existing_names.add(new_name)

        # --- Scale ---
        try:
            new_elev.Scale = section.Scale
        except Exception:
            pass

        # --- Phase ---
        try:
            phase_param = section.get_Parameter(BuiltInParameter.VIEW_PHASE)
            if phase_param and not phase_param.IsReadOnly:
                dst_phase = new_elev.get_Parameter(BuiltInParameter.VIEW_PHASE)
                if dst_phase and not dst_phase.IsReadOnly:
                    dst_phase.Set(phase_param.AsElementId())
        except Exception:
            pass

        # --- Far clip offset + enable "clip without line" ---
        try:
            src_far = section.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
            dst_far = new_elev.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
            if src_far and dst_far and not dst_far.IsReadOnly:
                dst_far.Set(src_far.AsDouble())
        except Exception:
            pass
        try:
            # 0 = No clip  1 = Clip without line  2 = Clip with line
            dst_fc = new_elev.get_Parameter(BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING)
            if dst_fc and not dst_fc.IsReadOnly:
                dst_fc.Set(1)
        except Exception:
            pass

        # --- Crop region (set after rotation + Regenerate) ---
        doc.Regenerate()
        _set_elevation_crop(section, new_elev)

        # --- View template (applied last so it wins on properties it controls) ---
        template_id = section.ViewTemplateId
        if template_id != ElementId.InvalidElementId:
            try:
                new_elev.ViewTemplateId = template_id
            except Exception:
                pass

        # --- Copy view-owned detail elements from section into new elevation ---
        # Copy one element at a time so a failing tag/dimension doesn't abort
        # the entire batch; detail lines, text, filled regions, etc. will succeed.
        elem_ids = _collect_view_owned_elements(doc, section)
        for src_id in elem_ids:
            try:
                copied = ElementTransformUtils.CopyElements(
                    section,
                    List[ElementId]([src_id]),
                    new_elev,
                    None,
                    copy_options,
                )
                if copied:
                    dest_id = list(copied)[0]
                    try:
                        new_elev.SetElementOverrides(
                            dest_id, section.GetElementOverrides(src_id)
                        )
                    except Exception:
                        pass
            except Exception:
                pass

        created.append((sec_name, new_name))

# ---------------------------------------------------------------------------
# Step 6 – Report
# ---------------------------------------------------------------------------

if not created and not skipped:
    forms.alert("Nothing was processed.")
    sys.exit(0)

lines = []

if created:
    lines.append("CREATED {} ELEVATION{}:".format(
        len(created), "S" if len(created) != 1 else ""
    ))
    for sec_name, elev_name in created:
        lines.append(u"  \u2022 {} \u2192 {}".format(sec_name, elev_name))

if skipped:
    lines.append("")
    lines.append("SKIPPED:")
    for sec_name, reason in skipped:
        lines.append(u"  \u2022 {} ({})".format(sec_name, reason))

forms.alert(
    "{} section{} converted to elevation{}.".format(
        len(created),
        "s" if len(created) != 1 else "",
        "s" if len(created) != 1 else "",
    ),
    expanded="\n".join(lines),
)
