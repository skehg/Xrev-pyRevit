# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from System.Collections.Generic import List
from pyrevit import revit, forms


doc = revit.doc
uidoc = revit.uidoc


def is_not_a_reference(ref_plane):
    # Using LookupParameter is more resilient across Revit API versions
    # where the related BuiltInParameter enum member may differ.
    param = ref_plane.LookupParameter("Is Reference")
    if not param:
        return False

    # Check exact match for "Not a Reference" (int=12 in this Revit build)
    val_string = (param.AsValueString() or param.AsString() or "").strip()
    return val_string.lower() == "not a reference"


def get_non_reference_planes():
    planes = (
        FilteredElementCollector(doc)
        .OfClass(ReferencePlane)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return [plane for plane in planes if is_not_a_reference(plane)]


def main():
    if not doc.IsFamilyDocument:
        forms.alert(
            "This tool only works in the Family Editor.",
            title="Select Non-Reference Planes",
            exitscript=True
        )

    non_reference_planes = get_non_reference_planes()
    if not non_reference_planes:
        forms.alert(
            "No reference planes found with Is Reference set to 'Not a Reference'.",
            title="Select Non-Reference Planes"
        )
        return

    element_ids = List[ElementId]([plane.Id for plane in non_reference_planes])
    uidoc.Selection.SetElementIds(element_ids)

    forms.alert(
        "Selected {} non-reference plane(s).".format(len(non_reference_planes)),
        title="Select Non-Reference Planes"
    )


main()
