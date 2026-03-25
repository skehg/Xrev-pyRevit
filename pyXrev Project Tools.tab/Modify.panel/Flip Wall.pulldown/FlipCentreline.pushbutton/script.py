# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, forms

uidoc = revit.uidoc
doc = revit.doc


# -----------------------------
# Selection Filter: Walls Only
# -----------------------------
class WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Wall)
    def AllowReference(self, ref, point):
        return True


# -----------------------------
# Flip Wall About True Centreline
# -----------------------------
def flip_about_centreline(wall):
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return

    # Store original location line
    original_loc_line = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).AsInteger()

    # Correct enum name
    centreline = int(WallLocationLine.WallCenterline)

    # Temporarily set to centreline
    wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set(centreline)

    # Flip
    wall.Flip()

    # Restore original location line
    wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set(original_loc_line)


# -----------------------------
# Main Continuous Flip Loop
# -----------------------------
def flip_walls_continuous():
    sel_filter = WallSelectionFilter()

    forms.toast(
        "Click walls to flip — press ESC to finish",
        title="Flip Walls",
        appid="flipwalls"
    )

    while True:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                sel_filter,
                "Click a wall to flip — press ESC to finish."
            )
            wall = doc.GetElement(ref.ElementId)

            with revit.Transaction("Flip Wall (Centreline)"):
                flip_about_centreline(wall)

        except Exception:
            # ESC ends the loop
            break


# -----------------------------
# Run
# -----------------------------
flip_walls_continuous()