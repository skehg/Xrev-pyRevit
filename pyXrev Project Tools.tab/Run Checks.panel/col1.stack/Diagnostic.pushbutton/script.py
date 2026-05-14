from Autodesk.Revit.DB import FilteredElementCollector, View

doc = __revit__.ActiveUIDocument.Document

# Grab one template
vt = FilteredElementCollector(doc).OfClass(View).ToElements()[0]
print(vt.IsTemplate)  # should be True
print(dir(vt))        # look for 'SetSubCategoryHidden' and 'IsSubCategoryHidden'