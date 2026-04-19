# -*- coding: utf-8 -*-
"""Set selected elements transparency in the active view."""
from pyrevit import revit, DB, forms, script as _pyscript
import sys


_ENV_KEY = "ObjectOverrides_TransparencySelected"

doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveGraphicalView
selection_ids = list(uidoc.Selection.GetElementIds())

if view is None:
    forms.alert("No active graphical view found.", exitscript=True)

if not selection_ids:
    forms.alert("Select one or more elements and try again.", exitscript=True)

saved_value = _pyscript.get_envvar(_ENV_KEY) or "100"
user_value = forms.ask_for_string(
    default=str(saved_value),
    prompt="Enter transparency value (0-100) for the selected elements.",
    title="Set Selected Elements Transparency",
)

if user_value is None:
    sys.exit(0)

user_value = user_value.strip()
if not user_value:
    sys.exit(0)

try:
    transparency = int(user_value)
    if transparency < 0 or transparency > 100:
        raise ValueError()
except ValueError:
    forms.alert("Transparency must be a whole number between 0 and 100.", exitscript=True)

_pyscript.set_envvar(_ENV_KEY, str(transparency))

updated = 0
skipped = 0

with revit.Transaction("Set Selected Elements Transparency"):
    for elem_id in selection_ids:
        element = doc.GetElement(elem_id)
        if element is None:
            skipped += 1
            continue

        try:
            overrides = view.GetElementOverrides(elem_id)
            overrides.SetSurfaceTransparency(transparency)
            view.SetElementOverrides(elem_id, overrides)
            updated += 1
        except Exception:
            skipped += 1

message = "Transparency {} applied to {} selected element(s).".format(transparency, updated)
if skipped:
    message += "\nSkipped {} element(s).".format(skipped)

forms.alert(message)
