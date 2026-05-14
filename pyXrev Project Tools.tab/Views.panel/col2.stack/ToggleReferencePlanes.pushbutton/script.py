# -*- coding: utf-8 -*-
"""Toggle Reference Plane visibility in the active view."""
from pyrevit import revit, DB, forms


doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveGraphicalView

if view is None:
    forms.alert("No active graphical view found.", exitscript=True)


def has_view_template(target_view):
    try:
        return target_view.ViewTemplateId != DB.ElementId.InvalidElementId
    except Exception:
        return False


def is_temp_view_props_enabled(target_view):
    try:
        return bool(target_view.IsTemporaryViewPropertiesModeEnabled())
    except Exception:
        pass

    try:
        return bool(
            target_view.IsTemporaryViewModeEnabled(
                DB.TemporaryViewMode.TemporaryViewProperties
            )
        )
    except Exception:
        return False


def get_reference_plane_category_id(document):
    for bic_name in ("OST_CLines", "OST_ReferencePlanes"):
        bic = getattr(DB.BuiltInCategory, bic_name, None)
        if bic is None:
            continue
        try:
            category = DB.Category.GetCategory(document, bic)
            if category:
                return category.Id
        except Exception:
            continue
    return None


def template_controls_annotations(template_view):
    if template_view is None:
        return False

    try:
        non_controlled_ids = set(
            [pid.IntegerValue for pid in template_view.GetNonControlledTemplateParameterIds()]
        )
    except Exception:
        non_controlled_ids = set()

    keywords = (
        "annotation",
        "visibility/graphics overrides",
        "visibility graphics overrides",
    )

    for param in template_view.Parameters:
        try:
            if param.Id.IntegerValue in non_controlled_ids:
                continue
            name = (param.Definition.Name or "").lower()
            if any(keyword in name for keyword in keywords):
                return True
        except Exception:
            continue

    return False


def toggle_reference_planes(target_view, ref_plane_category_id):
    is_hidden = target_view.GetCategoryHidden(ref_plane_category_id)
    target_view.SetCategoryHidden(ref_plane_category_id, not is_hidden)
    return "OFF" if not is_hidden else "ON"


category_id = get_reference_plane_category_id(doc)
if category_id is None:
    forms.alert("Could not find the Reference Planes category.", exitscript=True)

try:
    can_hide = view.CanCategoryBeHidden(category_id)
except Exception:
    can_hide = True

if not can_hide:
    forms.alert("Reference Planes cannot be toggled in this view. Enable Temporary View Properties and try again", exitscript=True)

view_has_template = has_view_template(view)
temp_mode_enabled = is_temp_view_props_enabled(view)
template = doc.GetElement(view.ViewTemplateId) if view_has_template else None
annotations_controlled = template_controls_annotations(template) if template else False

if view_has_template and annotations_controlled and not temp_mode_enabled:
    forms.alert(
        "This view template controls annotation visibility.\n\n"
        "Enter Temporary View Properties mode first, then run this tool again.",
        exitscript=True,
    )

try:
    with revit.Transaction("Toggle Reference Plane Visibility"):
        toggle_reference_planes(view, category_id)
except Exception as exc:
    forms.alert("Could not toggle Reference Planes visibility.\n{}".format(exc), exitscript=True)
