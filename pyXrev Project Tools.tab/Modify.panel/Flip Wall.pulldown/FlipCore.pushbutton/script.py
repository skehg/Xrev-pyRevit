# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, forms

uidoc = revit.uidoc
doc = revit.doc


class WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Wall)
    def AllowReference(self, ref, point):
        return True


def flip_about_core_centreline(wall):
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return

    loc_param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
    if not loc_param:
        return

    # Store original location line
    original_loc_line = loc_param.AsInteger()

    # --- STEP 1: Set location line to Core Centerline ---
    with revit.Transaction("Set Core Centreline"):
        loc_param.Set(int(WallLocationLine.CoreCenterline))

    # --- STEP 2: Flip (Revit now uses the new axis) ---
    with revit.Transaction("Flip Wall"):
        wall.Flip()

    # --- STEP 3: Restore original location line ---
    with revit.Transaction("Restore Location Line"):
        loc_param.Set(original_loc_line)


def flip_walls_continuous():
    sel_filter = WallSelectionFilter()

    forms.toast(
        "Click walls to flip about CORE centreline — ESC to finish",
        title="Flip Walls (Core Centreline)",
        appid="flipwalls-core"
    )

    while True:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                sel_filter,
                "Click a wall to flip about CORE centreline — ESC to finish."
            )
            wall = doc.GetElement(ref.ElementId)

            flip_about_core_centreline(wall)

        except Exception:
            break


flip_walls_continuous()