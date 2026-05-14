# -*- coding: utf-8 -*-
# pyRevit script (IronPython 2.7)
# Centreline detail lines with style mapping, split around door/window openings

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector, GraphicsStyle, GraphicsStyleType,
    Line, XYZ, BuiltInCategory, Transaction, Options, ViewDetailLevel,
    Solid, GeometryInstance, GeometryElement
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc


def normalize_select_result(sel):
    if sel is None:
        return None
    try:
        if hasattr(sel, '__iter__') and not isinstance(sel, basestring):
            return sel[0] if len(sel) > 0 else None
    except NameError:
        if hasattr(sel, '__iter__') and not isinstance(sel, str):
            return sel[0] if len(sel) > 0 else None
    return sel


def get_selected_walls():
    walls = []
    for id in uidoc.Selection.GetElementIds():
        el = doc.GetElement(id)
        if el and el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Walls):
            walls.append(el)
    return walls


def get_param_value_string(elem, name):
    p = elem.LookupParameter(name)
    if p is None:
        et = doc.GetElement(elem.GetTypeId())
        if et:
            p = et.LookupParameter(name)
    if p is None:
        return None

    for getter in (p.AsValueString, p.AsString):
        try:
            v = getter()
            if v:
                return v
        except:
            pass

    try:
        return str(p.AsInteger())
    except:
        pass

    try:
        return str(p.AsDouble())
    except:
        pass

    return None


def collect_wall_parameter_names(walls):
    names = set()
    for w in walls:
        for p in w.Parameters:
            try:
                names.add(p.Definition.Name)
            except:
                pass
        wt = doc.GetElement(w.GetTypeId())
        if wt:
            for p in wt.Parameters:
                try:
                    names.add(p.Definition.Name)
                except:
                    pass
    return sorted(names)


def collect_valid_line_styles():
    styles = []
    coll = FilteredElementCollector(doc).OfClass(GraphicsStyle)

    for gs in coll:
        try:
            if gs.GraphicsStyleType != GraphicsStyleType.Projection:
                continue
            cat = gs.GraphicsStyleCategory
            if not cat:
                continue
            if cat.Parent and cat.Parent.Id.IntegerValue == int(BuiltInCategory.OST_Lines):
                styles.append((gs.Name, gs))
            elif cat.Id.IntegerValue == int(BuiltInCategory.OST_Lines):
                styles.append((gs.Name, gs))
        except:
            pass

    uniq = {}
    for n, gs in styles:
        if n not in uniq:
            uniq[n] = gs

    return sorted(uniq.items(), key=lambda x: x[0].lower())


def get_opening_intervals_for_instance(inst, wall_curve, wall):
    """
    Get parameter intervals (0 to 1) along wall_curve where the instance creates an opening.
    Uses width and location point approach since solid intersection doesn't work.
    """
    intervals = []
    
    try:
        # Check if this is hosted on our wall
        if not (hasattr(inst, 'Host') and inst.Host and inst.Host.Id == wall.Id):
            return intervals
        
        # Get the family symbol (type)
        symbol = inst.Symbol if hasattr(inst, 'Symbol') else None
        if not symbol:
            return intervals
        
        # Get width parameter
        width_param = symbol.LookupParameter("Width")
        if not width_param:
            width_param = symbol.LookupParameter("Rough Width")
        
        if not width_param:
            return intervals
        
        width = width_param.AsDouble()
        if width <= 0:
            return intervals
        
        # Get instance location
        inst_loc = inst.Location
        if not (inst_loc and hasattr(inst_loc, 'Point')):
            return intervals
        
        pt = inst_loc.Point
        
        # Project point onto wall curve
        proj = wall_curve.Project(pt)
        if not proj:
            return intervals
        
        # CRITICAL FIX: The parameter from Project() is NOT normalized (0-1)
        # It's the actual distance along the curve. We need to normalize it.
        t_raw = proj.Parameter
        curve_length = wall_curve.Length
        
        # Normalize to 0-1 range
        t_center = t_raw / curve_length
        
        # Convert width to normalized parameter space
        width_normalized = width / curve_length
        
        # Calculate interval with small buffer (2% of opening width)
        buffer = width_normalized * 0.02
        t0 = t_center - (width_normalized / 2.0) - buffer
        t1 = t_center + (width_normalized / 2.0) + buffer
        
        # Clamp to valid range
        t0 = max(0.0, t0)
        t1 = min(1.0, t1)
        
        if t1 > t0:  # Valid interval
            intervals.append((t0, t1))
        
    except Exception as e:
        # Silent fail
        pass
    
    return intervals


def get_all_opening_intervals(wall, wall_curve):
    """Get all opening intervals for a wall"""
    intervals = []

    # Get all inserts (doors, windows, etc.)
    insert_ids = wall.FindInserts(True, True, True, True)
    for iid in insert_ids:
        inst = doc.GetElement(iid)
        if inst is None:
            continue
        cat = inst.Category
        if not cat:
            continue
        # Only process doors and windows
        if cat.Id.IntegerValue not in (
            int(BuiltInCategory.OST_Doors),
            int(BuiltInCategory.OST_Windows)
        ):
            continue

        inst_intervals = get_opening_intervals_for_instance(inst, wall_curve, wall)
        intervals.extend(inst_intervals)

    return intervals


def merge_intervals(intervals):
    """Merge overlapping intervals"""
    if not intervals:
        return []
    intervals = sorted(intervals, key=lambda x: x[0])
    merged = [intervals[0]]
    for (s, e) in intervals[1:]:
        ls, le = merged[-1]
        if s <= le:
            merged[-1] = (ls, max(le, e))
        else:
            merged.append((s, e))
    return merged


def subtract_intervals(full_start, full_end, cut_intervals):
    """
    Subtract cut intervals from the full range, returning remaining segments.
    Cut intervals should already be merged.
    """
    remaining = [(full_start, full_end)]
    for (c0, c1) in cut_intervals:
        new_remaining = []
        for (r0, r1) in remaining:
            # No overlap
            if c1 <= r0 or c0 >= r1:
                new_remaining.append((r0, r1))
                continue
            # Partial overlap - keep parts outside cut
            if c0 > r0:
                new_remaining.append((r0, c0))
            if c1 < r1:
                new_remaining.append((c1, r1))
        remaining = new_remaining
    # Filter out very small segments
    return [(a, b) for (a, b) in remaining if b - a > 0.001]


def pick_parameter(param_names):
    sel = forms.SelectFromList.show(param_names, title='Select wall property', multiselect=False)
    return normalize_select_result(sel)


def pick_style_for_value(value, style_names):
    sel = forms.SelectFromList.show(style_names, title='Line style for value: {}'.format(value), multiselect=False)
    return normalize_select_result(sel)


def main():
    walls = get_selected_walls()
    if not walls:
        TaskDialog.Show("Info", "Select walls first.")
        return

    param_names = collect_wall_parameter_names(walls)
    chosen_param = pick_parameter(param_names)
    if not chosen_param:
        return

    unique_vals = sorted(set(
        v for w in walls
        for v in [get_param_value_string(w, chosen_param)]
        if v not in (None, "")
    ))
    if not unique_vals:
        TaskDialog.Show("Info", "No values found for parameter '{}'.".format(chosen_param))
        return

    line_styles = collect_valid_line_styles()
    if not line_styles:
        TaskDialog.Show("Error", "No valid projection line styles found.")
        return

    style_names = [n for n, gs in line_styles]
    style_lookup = dict(line_styles)

    mapping = {}
    for val in unique_vals:
        sel = pick_style_for_value(val, style_names)
        if sel is None:
            TaskDialog.Show("Cancelled", "User cancelled.")
            return
        mapping[val] = style_lookup[sel]

    view = doc.ActiveView

    # GEOMETRY PHASE (no transaction): compute remaining intervals per wall
    wall_segments = {}  # wall.Id -> (curve, list of (t0, t1))
    debug_info = []
    
    for w in walls:
        loc = w.Location
        if not loc or not hasattr(loc, "Curve"):
            continue
        curve = loc.Curve
        if not curve:
            continue

        # Get all opening intervals
        cut_intervals = get_all_opening_intervals(w, curve)
        cut_intervals = merge_intervals(cut_intervals)
        
        # Subtract openings from full wall length
        remaining = subtract_intervals(0.0, 1.0, cut_intervals)
        
        wall_segments[w.Id] = (curve, remaining)
        
        # Debug info
        debug_info.append("Wall {}: {} openings, {} segments".format(
            w.Id.IntegerValue, 
            len(cut_intervals), 
            len(remaining)
        ))

    # Show debug info
 #   if debug_info:
 #       msg = "\n".join(debug_info)
 #       TaskDialog.Show("Opening Detection", msg)

    # CREATION PHASE (with transaction)
    t = Transaction(doc, "Place centreline detail lines (split at openings)")
    t.Start()
    created = 0

    for w in walls:
        val = get_param_value_string(w, chosen_param)
        if val not in mapping:
            continue
        if w.Id not in wall_segments:
            continue

        curve, segments = wall_segments[w.Id]

        # Get Z coordinate from view or wall
        try:
            z = view.Origin.Z
        except:
            z = curve.GetEndPoint(0).Z

        # Create detail line for each segment
        for (t0, t1) in segments:
            # Use normalized parameters (0-1) with Evaluate
            p0 = curve.Evaluate(t0, True)
            p1 = curve.Evaluate(t1, True)

            # Set Z to view level
            p0 = XYZ(p0.X, p0.Y, z)
            p1 = XYZ(p1.X, p1.Y, z)

            # Create the line
            ln = Line.CreateBound(p0, p1)
            detail = doc.Create.NewDetailCurve(view, ln)
            detail.LineStyle = mapping[val]
            created += 1

    t.Commit()
    TaskDialog.Show("Done", "Created {} detail line segments.".format(created))


main()