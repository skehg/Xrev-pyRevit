# -*- coding: utf-8 -*-
"""Scope Box Usage Report
Lists all Views, Levels, Grids, and Reference Planes using Scope Boxes.
Displays results with clickable element IDs for selection.
"""

__title__ = 'Scope Box\nUsage Report'
__author__ = 'Your Name'

from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms

# Get the current document
doc = revit.doc
output = script.get_output()

# Dictionary to store scope box usage
scope_box_usage = {}

# Get all scope boxes in the project
scope_boxes = DB.FilteredElementCollector(doc)\
    .OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Initialize dictionary with all scope boxes
for sb in scope_boxes:
    sb_name = sb.Name if sb.Name else "Unnamed Scope Box"
    scope_box_usage[sb.Id] = {
        'name': sb_name,
        'views': [],
        'levels': [],
        'grids': [],
        'ref_planes': []
    }

# Add entry for elements with no scope box
scope_box_usage['none'] = {
    'name': 'No Scope Box',
    'views': [],
    'levels': [],
    'grids': [],
    'ref_planes': []
}

# Check all views for scope box parameter
all_views = DB.FilteredElementCollector(doc)\
    .OfClass(DB.View)\
    .WhereElementIsNotElementType()\
    .ToElements()

for view in all_views:
    # Skip view templates
    if view.IsTemplate:
        continue
    
    # Get scope box parameter
    scope_box_param = view.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
    
    if scope_box_param and scope_box_param.HasValue:
        sb_id = scope_box_param.AsElementId()
        if sb_id and sb_id != DB.ElementId.InvalidElementId:
            if sb_id in scope_box_usage:
                scope_box_usage[sb_id]['views'].append({
                    'name': view.Name,
                    'id': view.Id,
                    'type': view.ViewType.ToString()
                })
        else:
            scope_box_usage['none']['views'].append({
                'name': view.Name,
                'id': view.Id,
                'type': view.ViewType.ToString()
            })

# Check Levels for scope box
levels = DB.FilteredElementCollector(doc)\
    .OfClass(DB.Level)\
    .WhereElementIsNotElementType()\
    .ToElements()

for level in levels:
    scope_box_param = level.get_Parameter(DB.BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
    
    if scope_box_param and scope_box_param.HasValue:
        sb_id = scope_box_param.AsElementId()
        if sb_id and sb_id != DB.ElementId.InvalidElementId:
            if sb_id in scope_box_usage:
                scope_box_usage[sb_id]['levels'].append({
                    'name': level.Name,
                    'id': level.Id
                })
        else:
            scope_box_usage['none']['levels'].append({
                'name': level.Name,
                'id': level.Id
            })
    else:
        scope_box_usage['none']['levels'].append({
            'name': level.Name,
            'id': level.Id
        })

# Check Grids for scope box
grids = DB.FilteredElementCollector(doc)\
    .OfClass(DB.Grid)\
    .WhereElementIsNotElementType()\
    .ToElements()

for grid in grids:
    scope_box_param = grid.get_Parameter(DB.BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
    
    if scope_box_param and scope_box_param.HasValue:
        sb_id = scope_box_param.AsElementId()
        if sb_id and sb_id != DB.ElementId.InvalidElementId:
            if sb_id in scope_box_usage:
                scope_box_usage[sb_id]['grids'].append({
                    'name': grid.Name,
                    'id': grid.Id
                })
        else:
            scope_box_usage['none']['grids'].append({
                'name': grid.Name,
                'id': grid.Id
            })
    else:
        scope_box_usage['none']['grids'].append({
            'name': grid.Name,
            'id': grid.Id
        })

# Check Reference Planes for scope box
ref_planes = DB.FilteredElementCollector(doc)\
    .OfClass(DB.ReferencePlane)\
    .WhereElementIsNotElementType()\
    .ToElements()

for ref_plane in ref_planes:
    scope_box_param = ref_plane.get_Parameter(DB.BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
    
    if scope_box_param and scope_box_param.HasValue:
        sb_id = scope_box_param.AsElementId()
        if sb_id and sb_id != DB.ElementId.InvalidElementId:
            if sb_id in scope_box_usage:
                scope_box_usage[sb_id]['ref_planes'].append({
                    'name': ref_plane.Name if ref_plane.Name else "Unnamed",
                    'id': ref_plane.Id
                })
        else:
            scope_box_usage['none']['ref_planes'].append({
                'name': ref_plane.Name if ref_plane.Name else "Unnamed",
                'id': ref_plane.Id
            })
    else:
        scope_box_usage['none']['ref_planes'].append({
            'name': ref_plane.Name if ref_plane.Name else "Unnamed",
            'id': ref_plane.Id
        })

# Output the results
output.print_md("# Scope Box Usage Report")
output.print_md("---")

# Process each scope box
for sb_id, data in sorted(scope_box_usage.items(), 
                          key=lambda x: x[1]['name']):
    
    # Skip scope boxes with no usage
    total_usage = (len(data['views']) + len(data['levels']) + 
                   len(data['grids']) + len(data['ref_planes']))
    
    if total_usage == 0:
        continue
    
    # Print scope box header
    if sb_id != 'none':
        output.print_md("## {}".format(data['name']))
        output.print_md("**Scope Box ID:** {}".format(
            output.linkify(sb_id)))
    else:
        output.print_md("## {}".format(data['name']))
    
    # Print Views
    if data['views']:
        output.print_md("### Views ({})".format(len(data['views'])))
        for view in sorted(data['views'], key=lambda x: x['name']):
            output.print_md("- {} - *{}* - ID: {}".format(
                view['name'],
                view['type'],
                output.linkify(view['id'])
            ))
    
    # Print Levels
    if data['levels']:
        output.print_md("### Levels ({})".format(len(data['levels'])))
        for level in sorted(data['levels'], key=lambda x: x['name']):
            output.print_md("- {} - ID: {}".format(
                level['name'],
                output.linkify(level['id'])
            ))
    
    # Print Grids
    if data['grids']:
        output.print_md("### Grids ({})".format(len(data['grids'])))
        for grid in sorted(data['grids'], key=lambda x: x['name']):
            output.print_md("- {} - ID: {}".format(
                grid['name'],
                output.linkify(grid['id'])
            ))
    
    # Print Reference Planes
    if data['ref_planes']:
        output.print_md("### Reference Planes ({})".format(len(data['ref_planes'])))
        for ref_plane in sorted(data['ref_planes'], key=lambda x: x['name']):
            output.print_md("- {} - ID: {}".format(
                ref_plane['name'],
                output.linkify(ref_plane['id'])
            ))
    
    output.print_md("---")

# Print summary
output.print_md("## Summary")
total_scope_boxes = len([k for k in scope_box_usage.keys() if k != 'none'])
output.print_md("**Total Scope Boxes in Project:** {}".format(total_scope_boxes))

used_scope_boxes = len([k for k, v in scope_box_usage.items() 
                        if k != 'none' and 
                        (v['views'] or v['levels'] or v['grids'] or v['ref_planes'])])
output.print_md("**Scope Boxes in Use:** {}".format(used_scope_boxes))
output.print_md("**Unused Scope Boxes:** {}".format(total_scope_boxes - used_scope_boxes))