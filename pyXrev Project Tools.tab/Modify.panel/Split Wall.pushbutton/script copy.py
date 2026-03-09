# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import ElementTransformUtils
from pyrevit import revit, forms
import clr
clr.AddReference("PresentationFramework")
from System.Windows.Controls import CheckBox, ComboBox, Button, ListBox, StackPanel, TextBlock

doc = revit.doc
uidoc = revit.uidoc


# ---------------------------------------------------------
# SELECTION FILTER FOR CATEGORIES
# ---------------------------------------------------------
class CategorySelectionFilter(ISelectionFilter):
    def __init__(self, allowed_cat_ids):
        self.allowed_cat_ids = allowed_cat_ids

    def AllowElement(self, element):
        cat = element.Category
        if cat and cat.Id in self.allowed_cat_ids:
            return True
        return False

    def AllowReference(self, reference, point):
        return False


# ---------------------------------------------------------
# CATEGORY + OPTIONS DIALOG
# ---------------------------------------------------------
class CategoryOptionsWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, 'split_categories_options.xaml')

    def get_result(self):
        cats = {
            BuiltInCategory.OST_Walls: self.chkWalls.IsChecked,
            BuiltInCategory.OST_Columns: self.chkColumns.IsChecked,
            BuiltInCategory.OST_StructuralColumns: self.chkStrColumns.IsChecked
        }
        allowed = [bic for bic, enabled in cats.items() if enabled]
        return {
            'categories': allowed,
            'lock': self.chkLock.IsChecked,
            'refmode': self.cmbRefMode.SelectedItem.Content
        }

    def on_select_all(self, sender, args):
        self.chkWalls.IsChecked = True
        self.chkColumns.IsChecked = True
        self.chkStrColumns.IsChecked = True

    def on_select_none(self, sender, args):
        self.chkWalls.IsChecked = False
        self.chkColumns.IsChecked = False
        self.chkStrColumns.IsChecked = False

    def on_ok(self, sender, args):
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


# ---------------------------------------------------------
# LEVEL SELECTION DIALOG
# ---------------------------------------------------------
class LevelSelectionWindow(forms.WPFWindow):
    def __init__(self, levels):
        forms.WPFWindow.__init__(self, 'split_levels.xaml')
        self._levels = sorted(levels, key=lambda l: l.Elevation)
        # populate listbox with checkboxes
        for lvl in self._levels:
            cb = CheckBox()
            cb.Content = lvl.Name
            cb.Tag = lvl
            cb.IsChecked = True
            self.lstLevels.Items.Add(cb)

    def get_selected_levels(self):
        result = []
        for item in self.lstLevels.Items:
            if item.IsChecked:
                result.append(item.Tag)
        return result

    def on_select_all(self, sender, args):
        for item in self.lstLevels.Items:
            item.IsChecked = True

    def on_select_none(self, sender, args):
        for item in self.lstLevels.Items:
            item.IsChecked = False

    def on_ok(self, sender, args):
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


# ---------------------------------------------------------
# GEOMETRY / INTERSECTION HELPERS
# ---------------------------------------------------------
def element_intersects_level(element, level):
    """Check if element's bounding box crosses the level elevation."""
    bbox = element.get_BoundingBox(None)
    if not bbox:
        print("NO BBOX:", element.Id)
        return False

    minz = bbox.Min.Z
    maxz = bbox.Max.Z
    lvlz = level.Elevation

    intersects = (minz < lvlz < maxz)

    if not intersects:
        print("SKIPPED:", element.Id.IntegerValue,
              "bboxZ=({:.3f}, {:.3f})".format(minz, maxz),
              "levelZ={:.3f}".format(lvlz))
    return intersects


from Autodesk.Revit.DB import WallUtils

from Autodesk.Revit.DB import WallUtils, Plane, Line, IntersectionResultArray, SetComparisonResult

def split_element_at_level(element, level):
    """Split a wall at the level elevation if it intersects."""
    if not isinstance(element, Wall):
        return []

    if not element_intersects_level(element, level):
        return []

    try:
        loc = element.Location
        if not isinstance(loc, LocationCurve):
            return []

        curve = loc.Curve
        z = level.Elevation

        # Create a horizontal plane at the level elevation
        plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, z))

        # Intersect the wall curve with the plane
        result = clr.Reference[IntersectionResultArray]()
        outcome = curve.Intersect(plane, result)

        if outcome != SetComparisonResult.Overlap:
            print("No intersection between wall {} and level {}".format(
                element.Id.IntegerValue, level.Name))
            return []

        # Extract the intersection point
        split_point = result.Item[0].XYZPoint

        # Perform the split
        new_wall_id = WallUtils.SplitWall(element, split_point)
        return [new_wall_id]

    except Exception as e:
        print("Split error on wall {} at level {}: {}".format(
            element.Id.IntegerValue, level.Name, e))
        return []


# ---------------------------------------------------------
# WALL REFERENCE HELPERS FOR LOCKING
# ---------------------------------------------------------
def get_wall_centerline_ref(wall):
    loc = wall.Location
    if isinstance(loc, LocationCurve):
        return loc.Curve.Reference
    return None


def get_wall_face_ref(wall, core=False):
    """Get a face reference for wall:
       core=False -> exterior face (Wall Face)
       core=True  -> core exterior face (Core Face)
    """
    try:
        if core:
            shell_type = ShellLayerType.CoreExterior
        else:
            shell_type = ShellLayerType.Exterior

        refs = HostObjectUtils.GetSideFaces(wall, shell_type)
        if refs and len(refs) > 0:
            return refs[0]
    except:
        pass
    return None


def get_wall_reference_for_mode(wall, mode):
    if mode == "Centerline":
        return get_wall_centerline_ref(wall)
    elif mode == "Core Face":
        return get_wall_face_ref(wall, core=True)
    elif mode == "Wall Face":
        return get_wall_face_ref(wall, core=False)
    else:
        return None


# ---------------------------------------------------------
# LOCKING FUNCTION
# ---------------------------------------------------------
def lock_segments(segment_ids, ref_mode):
    """Lock consecutive segments of the same wall using chosen reference."""
    for i in range(len(segment_ids) - 1):
        el1 = doc.GetElement(segment_ids[i])
        el2 = doc.GetElement(segment_ids[i + 1])

        # Only walls get reference-based locking here
        if not isinstance(el1, Wall) or not isinstance(el2, Wall):
            continue

        try:
            ref1 = get_wall_reference_for_mode(el1, ref_mode)
            ref2 = get_wall_reference_for_mode(el2, ref_mode)
            if not ref1 or not ref2:
                continue

            dim = doc.Create.NewAlignment(doc.ActiveView, ref1, ref2)
            dim.IsLocked = True

        except Exception as e:
            print("Lock error between {} and {}: {}".format(
                el1.Id.IntegerValue, el2.Id.IntegerValue, e))


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
def main():
    # 1. Category + options dialog
    cat_win = CategoryOptionsWindow()
    if not cat_win.show_dialog():
        return

    result = cat_win.get_result()
    bic_list = result['categories']
    lock_enabled = result['lock']
    ref_mode = result['refmode']

    if not bic_list:
        forms.alert("No categories selected. Operation cancelled.")
        return

    allowed_cat_ids = [ElementId(bic) for bic in bic_list]

    # 2. Persistent multi-view selection filtered by categories
    sel_filter = CategorySelectionFilter(allowed_cat_ids)
    selected_ids = set()

    while True:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                sel_filter,
                "Select elements to split (ESC to finish)"
            )
            if ref:
                selected_ids.add(ref.ElementId)
        except:
            # ESC or cancel ends selection
            break

    if not selected_ids:
        forms.alert("No elements selected. Operation cancelled.")
        return

    elements = [doc.GetElement(eid) for eid in selected_ids]

    # 3. Level selection
    all_levels = list(FilteredElementCollector(doc).OfClass(Level))
    if not all_levels:
        forms.alert("No levels found in document.")
        return

    lvl_win = LevelSelectionWindow(all_levels)
    if not lvl_win.show_dialog():
        return

    levels = lvl_win.get_selected_levels()
    if not levels:
        forms.alert("No levels selected. Operation cancelled.")
        return

    levels = sorted(levels, key=lambda l: l.Elevation)

    # 4. Split + optional locking
    with revit.Transaction("Split Elements by Selected Levels"):
        for el in elements:
            # Only process walls, columns, structural columns
            if isinstance(el, Wall):
                pass
            elif isinstance(el, FamilyInstance) and el.StructuralType == StructuralType.Column:
                pass
            else:
                continue

            segment_ids = [el.Id]

            for lvl in levels:
                new_ids = split_element_at_level(el, lvl)
                if new_ids:
                    segment_ids.extend(new_ids)

            if lock_enabled and len(segment_ids) > 1:
                # sort by Z so locking is consistent
                segment_ids = sorted(
                    segment_ids,
                    key=lambda id_:
                        doc.GetElement(id_).get_BoundingBox(None).Min.Z
                        if doc.GetElement(id_).get_BoundingBox(None) else 0.0
                )
                lock_segments(segment_ids, ref_mode)

    forms.alert("Splitting complete.")


if __name__ == "__main__":
    main()