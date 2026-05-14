# -*- coding: utf-8 -*-
"""Rename Family Types
Rename family types using pattern-based decomposition and reconstruction.
Mix extracted name parts with parameter values in any order.
"""

__title__ = 'Rename\nTypes'
__author__ = 'Xrev'

import clr
import os
import json
import re
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.Collections.ObjectModel import ObservableCollection
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, forms, coreutils

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

SCRIPT_DIR = os.path.dirname(__file__)
PARAM_CACHE_FILE = os.path.join(SCRIPT_DIR, 'param_cache.json')


def _safe_type_name(elem_type):
    """Best-effort type name getter across Revit API edge cases."""
    if elem_type is None:
        return ""

    try:
        name = revit.query.get_name(elem_type)
        if name:
            return str(name)
    except Exception:
        pass

    for bip in [BuiltInParameter.SYMBOL_NAME_PARAM,
                BuiltInParameter.ALL_MODEL_TYPE_NAME]:
        try:
            p = elem_type.get_Parameter(bip)
            if p and p.HasValue:
                val = p.AsString()
                if val:
                    return str(val)
        except Exception:
            pass

    return ""


def _safe_set_type_name(elem_type, new_name):
    """Best-effort type name setter across Revit API edge cases."""
    try:
        revit.update.set_name(elem_type, new_name)
        return True
    except Exception:
        pass

    for bip in [BuiltInParameter.SYMBOL_NAME_PARAM,
                BuiltInParameter.ALL_MODEL_TYPE_NAME]:
        try:
            p = elem_type.get_Parameter(bip)
            if p and not p.IsReadOnly:
                p.Set(new_name)
                return True
        except Exception:
            pass

    return False


def load_param_cache():
    """Load cached parameters from JSON file"""
    if os.path.exists(PARAM_CACHE_FILE):
        try:
            with open(PARAM_CACHE_FILE, 'r') as f:
                data = json.load(f)
                return data.get('parameters', []), data.get('selections', {})
        except:
            pass
    return [], {}


def save_param_cache(available_params, selections):
    """Save parameters and selections to JSON file"""
    try:
        data = {'parameters': available_params, 'selections': selections}
        with open(PARAM_CACHE_FILE, 'w') as f:
            json.dump(data, f, indent=2)
    except:
        pass


def collect_all_parameters(element_types):
    """Collect all unique string and numeric parameters from element types"""
    param_set = set()
    param_values_map = {}
    
    for elem_type in element_types:
        try:
            for param in elem_type.Parameters:
                try:
                    param_name = param.Definition.Name
                    param_storage = param.StorageType
                    
                    # Include String, Double (Length, Area, Angle), and Integer parameters
                    if param_storage in [StorageType.String, StorageType.Double, StorageType.Integer]:
                        param_set.add(param_name)
                        
                        # Store value if available
                        if param.HasValue:
                            try:
                                if param_storage == StorageType.String:
                                    val = param.AsString()
                                else:
                                    # For numeric types, use AsValueString to get formatted value with units
                                    val = param.AsValueString()
                                if val and param_name not in param_values_map:
                                    param_values_map[param_name] = str(val)
                            except:
                                pass
                except:
                    pass
        except:
            pass
    
    available_params = sorted(list(param_set))
    _, cached_selections = load_param_cache()
    save_param_cache(available_params, cached_selections)
    
    return available_params, param_values_map


# ============================================================================
# Load XAML from File
# ============================================================================
def load_xaml_window():
    """Load XAML window from TypeRenamerUI.xaml file"""
    script_dir = os.path.dirname(__file__)
    xaml_file = os.path.join(script_dir, 'TypeRenamerUI.xaml')
    
    if not os.path.exists(xaml_file):
        raise IOError("XAML file not found: {}".format(xaml_file))
    
    with open(xaml_file, 'r') as f:
        xaml_string = f.read()
    
    window = XamlReader.Parse(xaml_string)
    return window


# ============================================================================
# Preview Item Class
# ============================================================================
class PreviewItem(object):
    """Represents a single type rename preview"""
    def __init__(self, element_type, current_name):
        self.ElementType = element_type
        self.CurrentName = current_name
        self.NewName = current_name
        self.final = False
        self.tooltip = ''
    
    def format_value(self, old_pattern, new_pattern, param_values):
        """Format new name using pattern, only if not marked as final"""
        if self.final:
            return
        
        self.tooltip = ''
        try:
            if not old_pattern or not new_pattern:
                self.NewName = ''
                self.tooltip = 'Pattern Required'
                return
            
            # Clean the pattern (remove any trailing whitespace or special chars)
            old_pattern = old_pattern.strip()
            new_pattern = new_pattern.strip()
            
            # Convert template pattern {val1}-{val2}-{val3} to regex (.+?)-(.+?)-(.+)
            regex_pattern = re.sub(r'\{val\d+\}', '(.+?)', old_pattern)
            regex_pattern = '^' + regex_pattern + '$'  # Match entire string
            
            format_dict = {}
            
            try:
                # Use regex to extract groups
                match = re.match(regex_pattern, self.CurrentName)
                if match:
                    # Extract all groups and map them to val1, val2, etc.
                    groups = match.groups()
                    for i, group in enumerate(groups, 1):
                        format_dict['val{}'.format(i)] = group if group else ''
                else:
                    # Pattern didn't match - use empty values
                    for i in range(1, 11):
                        format_dict['val{}'.format(i)] = ''
            except Exception as e:
                # Regex error - use empty values
                for i in range(1, 11):
                    format_dict['val{}'.format(i)] = ''
            
            # Add parameter values to the format dictionary
            for i, val in enumerate(param_values, 1):
                format_dict['param{}'.format(i)] = val if val else ''
            
            # Apply the new pattern using string formatting
            try:
                self.NewName = new_pattern.format(**format_dict)
                self.tooltip = 'OK'
            except KeyError as ke:
                self.NewName = ''
                self.tooltip = 'Error: Unknown placeholder {}'.format(str(ke))
            
        except Exception as ex:
            self.NewName = ''
            self.tooltip = 'Error: {}'.format(str(ex))


# ============================================================================
# Rename Dialog Class
# ============================================================================
class RenameDialog(object):
    def __init__(self, element_types):
        """Initialize dialog with list of ElementTypes"""
        # Load XAML window
        self.window = load_xaml_window()
        
        # Store data
        self.element_types = element_types
        self.preview_items = ObservableCollection[object]()
        self.available_params = []
        self.param_values_map = {}
        self.result = False
        
        # Get controls - parameters (5 combos)
        self.param_combos = [
            self.window.FindName('combo_param1'),
            self.window.FindName('combo_param2'),
            self.window.FindName('combo_param3'),
            self.window.FindName('combo_param4'),
            self.window.FindName('combo_param5')
        ]
        
        # Get controls - patterns
        self.txt_original_pattern = self.window.FindName('txt_original_pattern')
        self.txt_new_pattern = self.window.FindName('txt_new_pattern')
        
        # Get controls - grid and buttons
        self.preview_grid = self.window.FindName('preview_grid')
        self.btn_apply = self.window.FindName('btn_apply')
        self.btn_cancel = self.window.FindName('btn_cancel')
        
        # Wire up events
        for combo in self.param_combos:
            combo.SelectionChanged += self.OnSettingsChanged
        
        self.txt_original_pattern.TextChanged += self.OnSettingsChanged
        self.txt_new_pattern.TextChanged += self.OnSettingsChanged
        self.preview_grid.SelectedCellsChanged += self.OnGridSelectionChanged
        
        self.btn_apply.Click += self.OnApply
        self.btn_cancel.Click += self.OnCancel
        
        # Initialize data
        self.initialize_data()
        
    def initialize_data(self):
        """Initialize parameter list and preview items"""
        # Collect all parameters and cache them
        self.available_params, self.param_values_map = collect_all_parameters(self.element_types)
        
        # Load previous selections from cache
        _, cached_selections = load_param_cache()
        
        # Populate parameter combo boxes (5 slots)
        for i, combo in enumerate(self.param_combos):
            if combo:
                combo.Items.Add('')
                for param_name in self.available_params:
                    combo.Items.Add(param_name)
                
                # Restore previous selection if available
                param_key = 'param{}'.format(i + 1)
                if param_key in cached_selections:
                    prev_selection = cached_selections[param_key]
                    try:
                        combo.SelectedItem = prev_selection
                    except:
                        combo.SelectedIndex = 0
        
        # Create preview items
        for elem_type in self.element_types:
            current_name = _safe_type_name(elem_type)
            if not current_name:
                current_name = "<Unnamed Type {}>".format(elem_type.Id.IntegerValue)
            
            preview_item = PreviewItem(elem_type, current_name)
            self.preview_items.Add(preview_item)
        
        # Set grid data source
        self.preview_grid.ItemsSource = self.preview_items
        
        # Set default patterns
        self.txt_original_pattern.Text = '{val1}'
        self.txt_new_pattern.Text = '{val1}'
    
    def get_param_value(self, elem_type, param_name):
        """Get parameter value as string from element type"""
        if not param_name:
            return ""
        
        # Ensure it's a string
        param_name = str(param_name).strip()
        if not param_name:
            return ""
        
        try:
            param = elem_type.LookupParameter(param_name)
            if param:
                if param.HasValue:
                    param_storage = param.StorageType
                    
                    if param_storage == StorageType.String:
                        value = param.AsString()
                        return str(value) if value else ""
                    elif param_storage == StorageType.Double:
                        # For Double (Length, Area, Angle, etc), use AsValueString for formatted value with units
                        value = param.AsValueString()
                        return str(value) if value else ""
                    elif param_storage == StorageType.Integer:
                        # For Integer types
                        value = param.AsInteger()
                        return str(value) if value is not None else ""
                    else:
                        # Try to get as string for other types
                        try:
                            value = param.AsString()
                            return str(value) if value else ""
                        except:
                            pass
        except Exception as e:
            print("Error getting param '{}': {}".format(param_name, str(e)))
        
        return ""
    
    def get_selected_param_values(self, elem_type):
        """Get values for all 5 selected parameters"""
        values = []
        for combo in self.param_combos:
            if combo:
                selected = combo.SelectedItem
                if selected:
                    param_name = str(selected).strip()
                    if param_name:
                        values.append(self.get_param_value(elem_type, param_name))
                    else:
                        values.append("")
                else:
                    values.append("")
            else:
                values.append("")
        return values
    
    def update_preview(self):
        """Update preview grid with new names using current patterns and parameters"""
        old_pattern = self.txt_original_pattern.Text
        new_pattern = self.txt_new_pattern.Text
        
        for item in self.preview_items:
            param_values = self.get_selected_param_values(item.ElementType)
            item.format_value(old_pattern, new_pattern, param_values)
        
        self.preview_grid.Items.Refresh()
    
    def OnSettingsChanged(self, sender, args):
        """Called when patterns or parameters change"""
        self.update_preview()
    
    def OnGridSelectionChanged(self, sender, args):
        """Called when grid selection changes (no preview update needed)"""
        pass
    
    def OnApply(self, sender, args):
        """Apply button clicked"""
        self.result = True
        self.window.DialogResult = True
        self.window.Close()
    
    def OnCancel(self, sender, args):
        """Cancel button clicked"""
        self.result = False
        self.window.DialogResult = False
        self.window.Close()
    
    def show_dialog(self):
        """Show the dialog and return result"""
        self.window.ShowDialog()
        return self.result
    
    def get_rename_map(self):
        """Get dictionary mapping ElementType to new name"""
        rename_map = {}
        for item in self.preview_items:
            if item.NewName and item.NewName != item.CurrentName:
                rename_map[item.ElementType] = item.NewName
        return rename_map


# ============================================================================
# Main Script Logic
# ============================================================================
def get_selected_types():
    """Get family types from current selection or allow user to pick"""
    selection = uidoc.Selection.GetElementIds()
    
    types = []
    
    # Try to get types from selection
    if selection.Count > 0:
        for elem_id in selection:
            try:
                elem = doc.GetElement(elem_id)
                if elem:
                    # If element is a type, use it directly
                    if isinstance(elem, ElementType):
                        types.append(elem)
                    # If element is an instance, get its type
                    else:
                        try:
                            type_id = elem.GetTypeId()
                            if type_id.Value != -1:  # Check for InvalidElementId
                                elem_type = doc.GetElement(type_id)
                                if elem_type and elem_type not in types:
                                    types.append(elem_type)
                        except:
                            pass
            except:
                pass
    
    # If no types found, let user pick types
    if len(types) == 0:
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements (their types will be renamed)")
            for ref in refs:
                try:
                    elem = doc.GetElement(ref.ElementId)
                    if elem:
                        if isinstance(elem, ElementType):
                            types.append(elem)
                        else:
                            try:
                                type_id = elem.GetTypeId()
                                if type_id.Value != -1:  # Check for InvalidElementId
                                    elem_type = doc.GetElement(type_id)
                                    if elem_type and elem_type not in types:
                                        types.append(elem_type)
                            except:
                                pass
                except:
                    pass
        except:
            # User cancelled
            return []
    
    return types


def rename_types(rename_map):
    """Rename types based on the rename map"""
    if not rename_map:
        forms.alert("No types to rename.", exitscript=True)
    
    t = Transaction(doc, "Rename Family Types")
    t.Start()
    
    try:
        success_count = 0
        failed_count = 0
        
        for elem_type, new_name in rename_map.items():
            try:
                if _safe_set_type_name(elem_type, new_name):
                    success_count += 1
                else:
                    print("ERROR renaming element id {} to '{}': No writable name field found.".format(
                        elem_type.Id.IntegerValue, new_name))
                    failed_count += 1
            except Exception as e:
                print("ERROR renaming element id {} to '{}': {}".format(
                    elem_type.Id.IntegerValue, new_name, str(e)))
                failed_count += 1
        
        t.Commit()
        
        message = "Renamed {} type(s) successfully.".format(success_count)
        if failed_count > 0:
            message += "\n{} type(s) failed (see output for details).".format(failed_count)
        
        forms.alert(message)
        
    except Exception as e:
        t.RollBack()
        forms.alert("Error during rename: {}".format(str(e)), exitscript=True)


# ============================================================================
# Main Execution
# ============================================================================
if __name__ == '__main__':
    # Get types to rename
    types = get_selected_types()
    
    if not types:
        forms.alert("No family types selected.", exitscript=True)
    
    # Show dialog
    dialog = RenameDialog(types)
    result = dialog.show_dialog()
    
    if result:
        # Get rename map
        rename_map = dialog.get_rename_map()
        
        if rename_map:
            # Perform rename
            rename_types(rename_map)
        else:
            forms.alert("No changes to apply.")