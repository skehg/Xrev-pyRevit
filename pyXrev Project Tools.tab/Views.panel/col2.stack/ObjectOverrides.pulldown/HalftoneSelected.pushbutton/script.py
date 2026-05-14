# -*- coding: utf-8 -*-
"""Set selected elements to halftone in the active view."""
from pyrevit import revit, DB, forms


doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveGraphicalView
selection_ids = list(uidoc.Selection.GetElementIds())

if view is None:
    forms.alert("No active graphical view found.", exitscript=True)

if not selection_ids:
    forms.alert("Select one or more elements and try again.", exitscript=True)

updated = 0
skipped = 0

with revit.Transaction("Set Selected Elements to Halftone"):
    for elem_id in selection_ids:
        element = doc.GetElement(elem_id)
        if element is None:
            skipped += 1
            continue

        try:
            overrides = view.GetElementOverrides(elem_id)
            overrides.SetHalftone(True)
            view.SetElementOverrides(elem_id, overrides)
            updated += 1
        except Exception:
            skipped += 1

message = "Halftone applied to {} selected element(s).".format(updated)
if skipped:
    message += "\nSkipped {} element(s).".format(skipped)

forms.alert(message)
