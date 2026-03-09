# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import ElementTransformUtils
from pyrevit import revit, forms
import clr

clr.AddReference("PresentationFramework")
from System.Windows.Controls import CheckBox

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
# (requires txtOffset TextBox in XAML for offset in mm)
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

        # Offset in mm (default 0 if empty or invalid)
        offset_mm = 0.0
        try:
            txt = self.txtOffset.Text.strip()
            if txt:
                offset_mm = float(txt)
        except:
            offset_mm = 0.0

        return {
            'categories': allowed,
            'lock': self.chkLock.IsChecked,
            'refmode': self.cmbRefMode.SelectedItem.Content,
            'offset_mm': offset_mm
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
# VERTICAL EXTENTS + INTERSECTION HELPERS
# ---------------------------------------------------------
def get_vertical_extents_and_params(element):
    """
    Return:
        base_z, top_z, base_lvl_id, base_off, top_lvl_id, top_off, kind
    kind = 'wall' or 'column' or None
    """
    kind = None

    if isinstance(element, Wall):
        kind = 'wall'
        base_lvl_param = BuiltInParameter.WALL_BASE_CONSTRAINT
        base_off_param = BuiltInParameter.WALL_BASE_OFFSET
        top_lvl_param = BuiltInParameter.WALL_HEIGHT_TYPE
        top_off_param = BuiltInParameter.WALL_TOP_OFFSET

    elif isinstance(element, FamilyInstance) and (
        element.StructuralType == StructuralType.Column
        or (element.Category and element.Category.Id.IntegerValue in
            [int(BuiltInCategory.OST_Columns), int(BuiltInCategory.OST_StructuralColumns)])
    ):
        kind = 'column'
        base_lvl_param = BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
        base_off_param = BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM
        top_lvl_param = BuiltInParameter.FAMILY_TOP_LEVEL_PARAM
        top_off_param = BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM

    if not kind:
        return None, None, None, None, None, None, None

    base_lvl_id = element.get_Parameter(base_lvl_param).AsElementId()
    top_lvl_id = element.get_Parameter(top_lvl_param).AsElementId()
    base_off = element.get_Parameter(base_off_param).AsDouble()
    top_off = element.get_Parameter(top_off_param).AsDouble()

    base_lvl = doc.GetElement(base_lvl_id)
    top_lvl = doc.GetElement(top_lvl_id)
    if not base_lvl or not top_lvl:
        return None, None, None, None, None, None, None

    base_z = base_lvl.Elevation + base_off
    top_z = top_lvl.Elevation + top_off

    return base_z, top_z, base_lvl_id, base_off, top_lvl_id, top_off, kind


def element_intersects_level(element, level, offset_ft):
    """Check if element's vertical extents cross the level elevation + offset."""
    base_z, top_z, _, _, _, _, _ = get_vertical_extents_and_params(element)
    if base_z is None:
        return False

    split_z = level.Elevation + offset_ft
    return base_z < split_z < top_z


# ---------------------------------------------------------
# MANUAL SPLITTING ENGINE (WALLS + COLUMNS)
# ---------------------------------------------------------
def split_vertical_element_by_levels(element, levels, offset_ft):
    """
    Manually split a vertical element (wall or column) by levels.
    - Original element becomes bottom-most segment.
    - New copies become upper segments.
    - Option B: top-most segment keeps original top constraint + offset.
    Returns list of segment Ids (bottom to top).
    """
    if not levels:
        return [element.Id]

    # Sort levels by elevation
    levels = sorted(levels, key=lambda l: l.Elevation)

    base_z, top_z, base_lvl_id, base_off, top_lvl_id, top_off, kind = \
        get_vertical_extents_and_params(element)

    if base_z is None or kind is None:
        return [element.Id]

    # Filter levels that actually intersect this element
    split_levels = [lvl for lvl in levels if element_intersects_level(element, lvl, offset_ft)]
    if not split_levels:
        return [element.Id]

    # Parameter ids per kind
    if kind == 'wall':
        base_lvl_param = BuiltInParameter.WALL_BASE_CONSTRAINT
        base_off_param = BuiltInParameter.WALL_BASE_OFFSET
        top_lvl_param = BuiltInParameter.WALL_HEIGHT_TYPE
        top_off_param = BuiltInParameter.WALL_TOP_OFFSET
    else:
        # columns
        base_lvl_param = BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
        base_off_param = BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM
        top_lvl_param = BuiltInParameter.FAMILY_TOP_LEVEL_PARAM
        top_off_param = BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM

    # Create N copies for N splits, total segments = N+1
    segment_ids = [element.Id]
    n = len(split_levels)
    for _ in split_levels:
        new_ids = ElementTransformUtils.CopyElement(doc, element.Id, XYZ(0, 0, 0))
        if new_ids and len(new_ids) > 0:
            segment_ids.append(new_ids[0])

    # Assign constraints to each segment
    # Seg0: base = original base, top = first split level + offset
    # Segi (1..n-1): base = split[i-1] + offset, top = split[i] + offset
    # SegN: base = split[n-1] + offset, top = original top + original top offset
    for idx, seg_id in enumerate(segment_ids):
        seg = doc.GetElement(seg_id)
        seg_base_lvl_param = seg.get_Parameter(base_lvl_param)
        seg_base_off_param = seg.get_Parameter(base_off_param)
        seg_top_lvl_param = seg.get_Parameter(top_lvl_param)
        seg_top_off_param = seg.get_Parameter(top_off_param)

        if idx == 0:
            # Bottom-most (original)
            seg_base_lvl_param.Set(base_lvl_id)
            seg_base_off_param.Set(base_off)
            seg_top_lvl_param.Set(split_levels[0].Id)
            seg_top_off_param.Set(offset_ft)
        elif idx == n:
            # Top-most
            seg_base_lvl_param.Set(split_levels[n - 1].Id)
            seg_base_off_param.Set(offset_ft)
            seg_top_lvl_param.Set(top_lvl_id)
            seg_top_off_param.Set(top_off)
        else:
            # Middle segments
            lower_lvl = split_levels[idx - 1]
            upper_lvl = split_levels[idx]
            seg_base_lvl_param.Set(lower_lvl.Id)
            seg_base_off_param.Set(offset_ft)
            seg_top_lvl_param.Set(upper_lvl.Id)
            seg_top_off_param.Set(offset_ft)

    return segment_ids


# ---------------------------------------------------------
# WALL REFERENCE HELPERS FOR LOCKING
# ---------------------------------------------------------
def get_wall_centerline_ref(wall):
    loc = wall.Location
    if isinstance(loc, LocationCurve):
        return loc.Curve.Reference
    return None


def get_wall_face_ref(wall, core=False):
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
    offset_mm = result['offset_mm']

    if not bic_list:
        forms.alert("No categories selected. Operation cancelled.")
        return

    # mm -> feet
    offset_ft = offset_mm / 304.8

    allowed_cat_ids = [ElementId(bic) for bic in bic_list]

    # 2. Selection with counter + ESC hint
    sel_filter = CategorySelectionFilter(allowed_cat_ids)
    selected_ids = set()

    forms.toast(
        "Select elements in any view — press ESC to finish",
        title="Split Elements",
        appid="split-elements"
    )

    while True:
        prompt = "Selected {} elements — press ESC to finish.".format(len(selected_ids))
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                sel_filter,
                prompt
            )
            if ref:
                selected_ids.add(ref.ElementId)
        except:
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
            is_wall = isinstance(el, Wall)
            is_col = isinstance(el, FamilyInstance) and (
                el.StructuralType == StructuralType.Column
                or (el.Category and el.Category.Id.IntegerValue in
                    [int(BuiltInCategory.OST_Columns), int(BuiltInCategory.OST_StructuralColumns)])
            )

            if not (is_wall or is_col):
                continue

            segment_ids = split_vertical_element_by_levels(el, levels, offset_ft)

            if lock_enabled and is_wall and len(segment_ids) > 1:
                # sort by base Z so locking is consistent
                def base_z_for_id(eid):
                    e = doc.GetElement(eid)
                    bz, _, _, _, _, _, _ = get_vertical_extents_and_params(e)
                    return bz if bz is not None else 0.0

                segment_ids = sorted(segment_ids, key=base_z_for_id)
                lock_segments(segment_ids, ref_mode)

    forms.alert("Splitting complete.")


if __name__ == "__main__":
    main()