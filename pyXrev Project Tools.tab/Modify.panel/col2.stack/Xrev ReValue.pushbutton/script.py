# -*- coding: utf-8 -*-
""""Xrev ReValue"
Re-value types/instances using advanced tokenization and pattern mapping.
Decompose names into tokens, apply parameter re-valuation, reconstruct with flexible syntax.
"""

__title__ = 'Xrev\nReValue'
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
HELP_EXAMPLES_FILE = os.path.join(SCRIPT_DIR, 'help_examples.md')

# ============================================================================
# DEFAULT TOKENIZATION CONFIG
# ============================================================================

DEFAULT_DELIMITERS = r'-_/,\|'
CANONICAL_DELIM = '|'

# Named delimiters for user-friendly syntax
NAMED_DELIMS = {
    'space': ' ',
    'hyphen': '-',
    'underscore': '_',
}


def build_tokenizer(delimiters):
    """Build a tokenizer function with the given delimiters."""
    try:
        # Escape delimiters for regex
        escaped = ''.join(['\\' + c if c in r'-[]{}()*+?.^\\|' else c for c in delimiters])
        delim_regex = re.compile(r'[%s]+' % escaped)
    except:
        # Fallback to default if regex fails
        delim_regex = re.compile(r'[%s]+' % DEFAULT_DELIMITERS)
    
    def tokenize(input_text):
        """Normalize and split input_text into tokens."""
        if not input_text:
            return []
        
        normalized = delim_regex.sub(CANONICAL_DELIM, input_text)
        raw_tokens = normalized.split(CANONICAL_DELIM)
        tokens = [t.strip() for t in raw_tokens if t.strip()]
        return tokens
    
    return tokenize


# Default tokenizer
tokenize = build_tokenizer(DEFAULT_DELIMITERS)

VAL_EXPR_REGEX = re.compile(r'\{([^{}]+)\}')

def _parse_val_expression(expr):
    """Parse the inside of {...}.
    
    Supported start specs:
      valN      - 1-based absolute index
      end       - last token
      end-N     - Nth token from the end (end-1 = 2nd to last, end-2 = 3rd to last)
    
    Supported forms (start[:end][|delim]):
      valN
      end
      end-N
      valN:M
      valN:+K
      valN:end
      valN:end-K
      end-N:end
      end-N:end-K
      end-N:+K
      valN:M|delim
      valN:end|delim
      end-N:end|delim
      end-N:end-K|delim
    
    Delimiter can be:
      - Named: space, hyphen, underscore
      - Literal: -, _, /, etc. (but NOT spaces as literals; use 'space' instead)
    
    Returns:
      (start_idx, end_spec, delimiter_string_or_None)
      start_idx is an int (1-based) or a string 'end'/'end-N'
    """
    # Split off optional delimiter part: "core|delim"
    if '|' in expr:
        core, delim_part = expr.split('|', 1)
        delim_part_stripped = delim_part.strip()
        # Check named delimiters first
        if delim_part_stripped in NAMED_DELIMS:
            delimiter = NAMED_DELIMS[delim_part_stripped]
        else:
            # Use stripped version as literal fallback (handles most cases)
            delimiter = delim_part_stripped if delim_part_stripped else None
    else:
        core = expr
        delimiter = None

    core = core.strip()

    def _parse_start(s):
        """Parse a start spec string into int (1-based) or string 'end'/'end-N'."""
        if s.startswith('val'):
            try:
                return int(s[3:])
            except ValueError:
                return None
        elif s == 'end':
            return 'end'
        elif s.startswith('end-'):
            try:
                int(s[4:])  # validate numeric suffix
                return s
            except ValueError:
                return None
        return None

    if ':' not in core:
        start_idx = _parse_start(core)
        if start_idx is None:
            return None, None, delimiter
        return start_idx, None, delimiter

    start_str, end_str = core.split(':', 1)
    start_str = start_str.strip()
    end_str = end_str.strip()

    start_idx = _parse_start(start_str)
    if start_idx is None:
        return None, None, delimiter

    if end_str == 'end':
        end_spec = 'end'
    elif end_str.startswith('end-'):
        # Handle "end-N" syntax (all except last N tokens)
        try:
            k = int(end_str[4:])
            end_spec = 'end-%d' % k
        except ValueError:
            return None, None, delimiter
    elif end_str.startswith('+'):
        try:
            k = int(end_str[1:])
        except ValueError:
            return None, None, delimiter
        end_spec = '+%d' % k
    else:
        try:
            end_idx = int(end_str)
        except ValueError:
            return None, None, delimiter
        end_spec = end_idx

    return start_idx, end_spec, delimiter


def _resolve_window(tokens, start_idx, end_spec):
    """Resolve a token window. start_idx may be int (1-based) or string 'end'/'end-N'."""
    n = len(tokens)
    if n == 0:
        return []

    # Resolve start_pos from start_idx
    if isinstance(start_idx, basestring):
        if start_idx == 'end':
            start_pos = n - 1
        elif start_idx.startswith('end-'):
            try:
                k = int(start_idx[4:])
            except ValueError:
                return []
            start_pos = n - 1 - k
            if start_pos < 0:
                start_pos = 0
        else:
            return []
    else:
        if start_idx < 1:
            start_idx = 1
        if start_idx > n:
            return []
        start_pos = start_idx - 1

    if end_spec is None:
        end_pos = start_pos
    elif end_spec == 'end':
        end_pos = n - 1
    elif isinstance(end_spec, basestring) and end_spec.startswith('end-'):
        try:
            k = int(end_spec[4:])
        except ValueError:
            return []
        end_pos = max(start_pos, n - 1 - k)
    elif isinstance(end_spec, int):
        end_pos = min(end_spec - 1, n - 1)
    elif isinstance(end_spec, basestring) and end_spec.startswith('+'):
        try:
            k = int(end_spec[1:])
        except ValueError:
            return []
        end_pos = min(start_pos + k, n - 1)
    else:
        return []

    if end_pos < start_pos:
        return []

    return tokens[start_pos:end_pos + 1]


def apply_pattern(input_text, pattern, param_dict, tokenizer_func=None):
    """Apply pattern to input_text using tokenization and parameter substitution.
    
    First applies tokenization for {valN}, then parameter substitution for {paramN}.
    """
    if tokenizer_func is None:
        tokenizer_func = tokenize
    
    tokens = tokenizer_func(input_text)
    
    def replace_match(m):
        inner = m.group(1).strip()
        
        # Try parameter substitution first ({paramN})
        if inner.startswith('param'):
            try:
                param_num = int(inner[5:])
                key = 'param{}'.format(param_num)
                if key in param_dict:
                    return param_dict[key]
                return ''
            except ValueError:
                pass
        
        # Try token extraction ({valN}, {valN:M}, etc)
        start_idx, end_spec, delim = _parse_val_expression(inner)

        if start_idx is None:
            return m.group(0)

        window_tokens = _resolve_window(tokens, start_idx, end_spec)
        joiner = delim if delim is not None else '-'
        return joiner.join(window_tokens)

    result = VAL_EXPR_REGEX.sub(replace_match, pattern)
    return result


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

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


def _safe_element_name(element):
    """Get name of element (instance), fallback to type."""
    if element is None:
        return ""
    
    try:
        name = revit.query.get_name(element)
        if name:
            return str(name)
    except Exception:
        pass
    
    return ""


def _safe_set_element_name(element, new_name):
    """Set name of element (instance)."""
    if element is None:
        return False
    
    try:
        element.Name = new_name
        return True
    except Exception:
        pass

    try:
        revit.update.set_name(element, new_name)
        return True
    except Exception:
        pass
    
    return False


def _find_parameter_for_write(element, param_name, is_instance_mode):
    """Find writable parameter by mode. Returns parameter or None."""
    if element is None or not param_name:
        return None

    param_name = str(param_name).strip()
    if not param_name:
        return None

    if is_instance_mode:
        # In instance mode, prefer instance parameter first.
        try:
            p = element.LookupParameter(param_name)
            if p:
                return p
        except Exception:
            pass

        elem_type = _get_element_type(element)
        if elem_type:
            try:
                p = elem_type.LookupParameter(param_name)
                if p:
                    return p
            except Exception:
                pass
    else:
        # Type mode writes only to type parameter.
        elem_type = _get_element_type(element)
        if elem_type:
            try:
                p = elem_type.LookupParameter(param_name)
                if p:
                    return p
            except Exception:
                pass

    return None


def _safe_set_parameter_value(element, param_name, new_value, is_instance_mode):
    """Set parameter value with basic storage-type handling."""
    p = _find_parameter_for_write(element, param_name, is_instance_mode)
    if not p or p.IsReadOnly:
        return False

    try:
        storage = p.StorageType
        text_value = '' if new_value is None else str(new_value)

        if storage == StorageType.String:
            p.Set(text_value)
            return True

        if storage in [StorageType.Double, StorageType.Integer]:
            # Try formatted string first (units-aware for doubles), then numeric parse.
            try:
                if p.SetValueString(text_value):
                    return True
            except Exception:
                pass

            try:
                if storage == StorageType.Integer:
                    p.Set(int(float(text_value)))
                else:
                    p.Set(float(text_value))
                return True
            except Exception:
                return False

    except Exception:
        return False

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


def _get_element_type(element):
    """Get ElementType for an element, or return the element if already a type."""
    if element is None:
        return None

    if isinstance(element, ElementType):
        return element

    try:
        type_id = element.GetTypeId()
        if type_id and type_id.Value != -1:
            return doc.GetElement(type_id)
    except:
        pass

    return None


def collect_all_parameters(elements, is_instance_mode):
    """Collect all unique string and numeric parameters from elements"""
    param_set = set()
    param_values_map = {}
    
    for elem in elements:
        try:
            targets = []

            if is_instance_mode:
                # Instance mode can use both instance and type parameters.
                targets.append(elem)
                elem_type = _get_element_type(elem)
                if elem_type:
                    targets.append(elem_type)
            else:
                # Type mode must only use type parameters.
                elem_type = _get_element_type(elem)
                if elem_type:
                    targets.append(elem_type)

            for target in targets:
                if target is None:
                    continue

                for param in target.Parameters:
                    try:
                        param_name = param.Definition.Name
                        param_storage = param.StorageType
                        
                        if param_storage in [StorageType.String, StorageType.Double, StorageType.Integer]:
                            param_set.add(param_name)
                            
                            if param.HasValue:
                                try:
                                    if param_storage == StorageType.String:
                                        val = param.AsString()
                                    else:
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
    """Load XAML window from ReValueUI.xaml file"""
    script_dir = os.path.dirname(__file__)
    xaml_file = os.path.join(script_dir, 'ReValueUI.xaml')
    
    if not os.path.exists(xaml_file):
        raise IOError("XAML file not found: {}".format(xaml_file))
    
    with open(xaml_file, 'r') as f:
        xaml_string = f.read()
    
    window = XamlReader.Parse(xaml_string)
    return window


def load_help_examples_text():
    """Load syntax/examples content from markdown file in this folder."""
    default_text = (
        "Xrev ReValue - Syntax and Examples\n\n"
        "Edit help_examples.md in this folder to customize this text."
    )

    if os.path.exists(HELP_EXAMPLES_FILE):
        try:
            with open(HELP_EXAMPLES_FILE, 'r') as f:
                return f.read()
        except Exception as ex:
            return "{}\n\nError reading help_examples.md: {}".format(default_text, str(ex))

    return default_text


# ============================================================================
# Preview Item Class
# ============================================================================
class PreviewItem(object):
    """Represents a single element re-value preview"""
    def __init__(self, element, current_name):
        self.Element = element
        self.CurrentName = current_name
        self.NewName = current_name
        self.final = False
        self.tooltip = ''
    
    def format_value(self, pattern, param_values, tokenizer_func):
        """Format new name using tokenization and parameter substitution"""
        if self.final:
            return
        
        self.tooltip = ''
        try:
            if not pattern:
                self.NewName = ''
                self.tooltip = 'Pattern Required'
                return
            
            pattern = pattern.strip()
            
            # Build parameter dict for substitution
            param_dict = {}
            for i, val in enumerate(param_values, 1):
                param_dict['param{}'.format(i)] = val if val else ''
            
            # Use the tokenization + parameter substitution pattern application
            self.NewName = apply_pattern(self.CurrentName, pattern, param_dict, tokenizer_func)
            self.tooltip = 'OK'
            
        except Exception as ex:
            self.NewName = ''
            self.tooltip = 'Error: {}'.format(str(ex))


# ============================================================================
# Re-Value Dialog Class
# ============================================================================
class ReValueDialog(object):
    def __init__(self, elements, is_instance_mode=False):
        """Initialize dialog with list of elements"""
        # Load XAML window
        self.window = load_xaml_window()
        
        # Store data
        self.elements = elements
        self.is_instance_mode = is_instance_mode
        self.preview_items = ObservableCollection[object]()
        self.available_params = []
        self.param_values_map = {}
        self.result = False
        
        # Get controls - mode toggle
        self.radio_type = self.window.FindName('radio_type')
        self.radio_instance = self.window.FindName('radio_instance')
        
        # Get controls - delimiters
        self.txt_delimiters = self.window.FindName('txt_delimiters')
        
        # Get controls - original parameter dropdown
        self.combo_orig_param = self.window.FindName('combo_orig_param')
        
        # Get controls - parameters (5 combos)
        self.param_combos = [
            self.window.FindName('combo_param1'),
            self.window.FindName('combo_param2'),
            self.window.FindName('combo_param3'),
            self.window.FindName('combo_param4'),
            self.window.FindName('combo_param5')
        ]
        
        # Get controls - pattern
        self.txt_pattern = self.window.FindName('txt_pattern')
        
        # Get controls - tokenization display
        self.txt_tokenization = self.window.FindName('txt_tokenization')
        
        # Get controls - grid and buttons
        self.preview_grid = self.window.FindName('preview_grid')
        self.btn_syntax_examples = self.window.FindName('btn_syntax_examples')
        self.btn_apply = self.window.FindName('btn_apply')
        self.btn_cancel = self.window.FindName('btn_cancel')
        
        # Wire up events
        if self.radio_type:
            self.radio_type.Checked += self.OnModeChanged
        if self.radio_instance:
            self.radio_instance.Checked += self.OnModeChanged
        if self.txt_delimiters:
            self.txt_delimiters.TextChanged += self.OnSettingsChanged
        if self.combo_orig_param:
            self.combo_orig_param.SelectionChanged += self.OnSettingsChanged
        
        for combo in self.param_combos:
            combo.SelectionChanged += self.OnSettingsChanged
        
        self.txt_pattern.TextChanged += self.OnSettingsChanged
        self.preview_grid.SelectedCellsChanged += self.OnGridSelectionChanged

        if self.btn_syntax_examples:
            self.btn_syntax_examples.Click += self.OnSyntaxExamples
        
        self.btn_apply.Click += self.OnApply
        self.btn_cancel.Click += self.OnCancel
        
        # Initialize data
        self.initialize_data()
        
    def initialize_data(self):
        """Initialize parameter list and preview items"""
        # Set mode
        if self.radio_type:
            self.radio_type.IsChecked = not self.is_instance_mode
        if self.radio_instance:
            self.radio_instance.IsChecked = self.is_instance_mode
        
        # Set default delimiters
        if self.txt_delimiters:
            self.txt_delimiters.Text = DEFAULT_DELIMITERS
        
        # Collect all parameters and cache them
        self.available_params, self.param_values_map = collect_all_parameters(
            self.elements, self.is_instance_mode)
        
        # Load previous selections from cache
        _, cached_selections = load_param_cache()
        
        # Populate original parameter combo box
        if self.combo_orig_param:
            self.combo_orig_param.Items.Add('')
            for param_name in self.available_params:
                self.combo_orig_param.Items.Add(param_name)
            
            param_key = 'orig_param'
            if param_key in cached_selections:
                prev_selection = cached_selections[param_key]
                try:
                    self.combo_orig_param.SelectedItem = prev_selection
                except:
                    self.combo_orig_param.SelectedIndex = 0
        
        # Populate parameter combo boxes (5 slots)
        for i, combo in enumerate(self.param_combos):
            if combo:
                combo.Items.Add('')
                for param_name in self.available_params:
                    combo.Items.Add(param_name)
                
                param_key = 'param{}'.format(i + 1)
                if param_key in cached_selections:
                    prev_selection = cached_selections[param_key]
                    try:
                        combo.SelectedItem = prev_selection
                    except:
                        combo.SelectedIndex = 0
        
        # Create preview items
        for elem in self.elements:
            if self.is_instance_mode:
                current_name = _safe_element_name(elem)
                if not current_name:
                    try:
                        current_name = "<Instance {}>".format(elem.Id.IntegerValue)
                    except:
                        current_name = "<Unnamed Instance>"
            else:
                # Type mode - get the type
                elem_type = None
                if isinstance(elem, ElementType):
                    elem_type = elem
                else:
                    try:
                        type_id = elem.GetTypeId()
                        if type_id.Value != -1:
                            elem_type = doc.GetElement(type_id)
                    except:
                        pass
                
                if elem_type:
                    current_name = _safe_type_name(elem_type)
                    if not current_name:
                        current_name = "<Unnamed Type {}>".format(elem_type.Id.IntegerValue)
                    elem = elem_type
                else:
                    continue
            
            preview_item = PreviewItem(elem, current_name)
            self.preview_items.Add(preview_item)
        
        # Set grid data source
        self.preview_grid.ItemsSource = self.preview_items
        
        # Set default pattern
        self.txt_pattern.Text = '{val1}'
        
        # Update tokenization display for first item
        self.update_tokenization_display()
    
    def get_delimiters(self):
        """Get current delimiters from UI"""
        if self.txt_delimiters:
            return self.txt_delimiters.Text or DEFAULT_DELIMITERS
        return DEFAULT_DELIMITERS
    
    def get_tokenizer_func(self):
        """Get tokenizer function with current delimiters"""
        delimiters = self.get_delimiters()
        return build_tokenizer(delimiters)
    
    def get_param_value(self, element, param_name):
        """Get parameter value as string from element"""
        if not param_name:
            return ""
        
        param_name = str(param_name).strip()
        if not param_name:
            return ""
        
        try:
            # In type mode, resolve selected element to type before lookup.
            target = element
            if not self.is_instance_mode:
                resolved_type = _get_element_type(element)
                if resolved_type:
                    target = resolved_type

            # Instance mode may include type parameters in selectors.
            # Try instance first, then fallback to type.
            if self.is_instance_mode:
                param = None
                try:
                    param = target.LookupParameter(param_name)
                except:
                    param = None

                if not param:
                    resolved_type = _get_element_type(element)
                    if resolved_type:
                        try:
                            param = resolved_type.LookupParameter(param_name)
                        except:
                            param = None
            else:
                param = target.LookupParameter(param_name)

            if param:
                if param.HasValue:
                    param_storage = param.StorageType
                    
                    if param_storage == StorageType.String:
                        value = param.AsString()
                        return str(value) if value else ""
                    elif param_storage == StorageType.Double:
                        value = param.AsValueString()
                        return str(value) if value else ""
                    elif param_storage == StorageType.Integer:
                        value = param.AsInteger()
                        return str(value) if value is not None else ""
                    else:
                        try:
                            value = param.AsString()
                            return str(value) if value else ""
                        except:
                            pass
        except Exception as e:
            print("Error getting param '{}': {}".format(param_name, str(e)))
        
        return ""
    
    def get_selected_param_values(self, element):
        """Get values for all 5 selected parameters"""
        values = []
        for combo in self.param_combos:
            if combo:
                selected = combo.SelectedItem
                if selected:
                    param_name = str(selected).strip()
                    if param_name:
                        values.append(self.get_param_value(element, param_name))
                    else:
                        values.append("")
                else:
                    values.append("")
            else:
                values.append("")
        return values
    
    def get_original_source_text(self, element):
        """Get the text to use as source for tokenization"""
        if self.combo_orig_param:
            selected = self.combo_orig_param.SelectedItem
            if selected:
                param_name = str(selected).strip()
                if param_name:
                    return self.get_param_value(element, param_name)
        
        # Fall back to element/type name
        if self.is_instance_mode:
            return _safe_element_name(element)
        else:
            return _safe_type_name(element)

    def get_selected_original_param_name(self):
        """Get selected original/source parameter name, or empty string for Name mode."""
        if not self.combo_orig_param:
            return ''

        selected = self.combo_orig_param.SelectedItem
        if not selected:
            return ''

        return str(selected).strip()
    
    def update_tokenization_display(self):
        """Update tokenization display for first selected item"""
        if self.txt_tokenization and len(self.preview_items) > 0:
            first_item = self.preview_items[0]
            source_text = self.get_original_source_text(first_item.Element)
            
            tokenizer_func = self.get_tokenizer_func()
            tokens = tokenizer_func(source_text)
            
            if tokens:
                token_display = ', '.join(['[{0}] {1}'.format(i+1, t) for i, t in enumerate(tokens)])
                self.txt_tokenization.Text = 'Tokens: {}'.format(token_display)
            else:
                self.txt_tokenization.Text = 'Tokens: (none)'
    
    def update_preview(self):
        """Update preview grid with new names using current pattern"""
        pattern = self.txt_pattern.Text
        tokenizer_func = self.get_tokenizer_func()
        
        for item in self.preview_items:
            # CurrentName column always reflects the selected tokenization source.
            item.CurrentName = self.get_original_source_text(item.Element)
            param_values = self.get_selected_param_values(item.Element)
            item.format_value(pattern, param_values, tokenizer_func)
        
        self.preview_grid.Items.Refresh()
        self.update_tokenization_display()
    
    def OnModeChanged(self, sender, args):
        """Called when instance/type mode changes"""
        # Re-initialize with new mode
        if self.radio_instance:
            self.is_instance_mode = self.radio_instance.IsChecked
        else:
            self.is_instance_mode = False
        
        # Recollect parameters and reinitialize
        self.available_params, self.param_values_map = collect_all_parameters(
            self.elements, self.is_instance_mode)

        _, cached_selections = load_param_cache()

        # Clear and rebuild original parameter combo
        if self.combo_orig_param:
            self.combo_orig_param.Items.Clear()
            self.combo_orig_param.Items.Add('')
            for param_name in self.available_params:
                self.combo_orig_param.Items.Add(param_name)

            prev_orig = cached_selections.get('orig_param', '')
            if prev_orig and prev_orig in self.available_params:
                self.combo_orig_param.SelectedItem = prev_orig
            else:
                self.combo_orig_param.SelectedIndex = 0
        
        # Clear and rebuild parameter combos
        for i, combo in enumerate(self.param_combos):
            combo.Items.Clear()
            combo.Items.Add('')
            for param_name in self.available_params:
                combo.Items.Add(param_name)

            param_key = 'param{}'.format(i + 1)
            prev_selection = cached_selections.get(param_key, '')
            if prev_selection and prev_selection in self.available_params:
                combo.SelectedItem = prev_selection
            else:
                combo.SelectedIndex = 0
        
        # Rebuild preview items
        self.preview_items.Clear()
        for elem in self.elements:
            if self.is_instance_mode:
                current_name = _safe_element_name(elem)
                if not current_name:
                    try:
                        current_name = "<Instance {}>".format(elem.Id.IntegerValue)
                    except:
                        current_name = "<Unnamed Instance>"
            else:
                elem_type = None
                if isinstance(elem, ElementType):
                    elem_type = elem
                else:
                    try:
                        type_id = elem.GetTypeId()
                        if type_id.Value != -1:
                            elem_type = doc.GetElement(type_id)
                    except:
                        pass
                
                if elem_type:
                    current_name = _safe_type_name(elem_type)
                    if not current_name:
                        current_name = "<Unnamed Type {}>".format(elem_type.Id.IntegerValue)
                    elem = elem_type
                else:
                    continue
            
            preview_item = PreviewItem(elem, current_name)
            self.preview_items.Add(preview_item)
        
        self.preview_grid.Items.Refresh()
        self.update_preview()
    
    def OnSettingsChanged(self, sender, args):
        """Called when patterns or parameters change"""
        self.update_preview()
    
    def OnGridSelectionChanged(self, sender, args):
        """Called when grid selection changes"""
        self.update_tokenization_display()

    def OnSyntaxExamples(self, sender, args):
        """Open temporary window with syntax and examples from help_examples.md."""
        help_text = load_help_examples_text()

        help_xaml = """
<Window xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"
        xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"
        Title=\"Syntax and Examples\"
        Height=\"700\" Width=\"900\"
        WindowStartupLocation=\"CenterOwner\"
        ResizeMode=\"CanResize\">
    <Grid Margin=\"12\">
        <Grid.RowDefinitions>
            <RowDefinition Height=\"*\"/>
            <RowDefinition Height=\"Auto\"/>
        </Grid.RowDefinitions>
        <TextBox x:Name=\"txt_help\"
                 Grid.Row=\"0\"
                 IsReadOnly=\"True\"
                 TextWrapping=\"Wrap\"
                 AcceptsReturn=\"True\"
                 VerticalScrollBarVisibility=\"Auto\"
                 HorizontalScrollBarVisibility=\"Auto\"
                 FontFamily=\"Consolas\"
                 FontSize=\"12\"
                 Background=\"WhiteSmoke\"
                 Padding=\"8\"/>
        <StackPanel Grid.Row=\"1\" Orientation=\"Horizontal\" HorizontalAlignment=\"Right\" Margin=\"0,10,0,0\">
            <Button x:Name=\"btn_close_help\" Content=\"Close\" Width=\"90\" Height=\"28\"/>
        </StackPanel>
    </Grid>
</Window>
"""

        try:
            help_window = XamlReader.Parse(help_xaml)
            txt_help = help_window.FindName('txt_help')
            btn_close_help = help_window.FindName('btn_close_help')

            if txt_help:
                txt_help.Text = help_text

            if btn_close_help:
                def _close_help_window(s, e):
                    help_window.Close()
                btn_close_help.Click += _close_help_window

            help_window.Owner = self.window
            help_window.ShowDialog()
        except Exception as ex:
            forms.alert("Could not open Syntax and Examples window: {}".format(str(ex)))
    
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
    
    def get_revalue_map(self):
        """Get dictionary mapping element to new name"""
        revalue_map = {}
        for item in self.preview_items:
            if item.NewName and item.NewName != item.CurrentName:
                revalue_map[item.Element] = item.NewName
        return revalue_map


# ============================================================================
# Main Script Logic
# ============================================================================
def get_selected_elements():
    """Get elements from current selection or allow user to pick"""
    selection = uidoc.Selection.GetElementIds()
    
    elements = []
    
    if selection.Count > 0:
        for elem_id in selection:
            try:
                elem = doc.GetElement(elem_id)
                if elem:
                    elements.append(elem)
            except:
                pass
    
    if len(elements) == 0:
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements to re-value")
            for ref in refs:
                try:
                    elem = doc.GetElement(ref.ElementId)
                    if elem and elem not in elements:
                        elements.append(elem)
                except:
                    pass
        except:
            return []
    
    return elements


def revalue_elements(revalue_map, is_instance_mode, target_param_name=''):
    """Re-value elements based on the revalue map (name or parameter target)."""
    if not revalue_map:
        forms.alert("No elements to re-value.", exitscript=True)

    try:
        success_count = 0
        failed_count = 0

        with revit.Transaction("Re-Value Elements"):
            for element, new_name in revalue_map.items():
                try:
                    if target_param_name:
                        success = _safe_set_parameter_value(
                            element, target_param_name, new_name, is_instance_mode)
                    else:
                        if is_instance_mode:
                            success = _safe_set_element_name(element, new_name)
                        else:
                            success = _safe_set_type_name(element, new_name)

                    if success:
                        success_count += 1
                    else:
                        if target_param_name:
                            print("ERROR re-valuing element id {} parameter '{}' to '{}': No writable parameter found.".format(
                                element.Id.IntegerValue, target_param_name, new_name))
                        else:
                            print("ERROR re-valuing element id {} to '{}': No writable name field found.".format(
                                element.Id.IntegerValue, new_name))
                        failed_count += 1
                except Exception as e:
                    if target_param_name:
                        print("ERROR re-valuing element id {} parameter '{}' to '{}': {}".format(
                            element.Id.IntegerValue, target_param_name, new_name, str(e)))
                    else:
                        print("ERROR re-valuing element id {} to '{}': {}".format(
                            element.Id.IntegerValue, new_name, str(e)))
                    failed_count += 1
        
        message = "Re-valued {} element(s) successfully.".format(success_count)
        if failed_count > 0:
            message += "\n{} element(s) failed (see output for details).".format(failed_count)
        
        forms.alert(message)
        
    except Exception as e:
        forms.alert("Error during re-value: {}".format(str(e)), exitscript=True)


# ============================================================================
# Main Execution
# ============================================================================
if __name__ == '__main__':
    elements = get_selected_elements()
    
    if not elements:
        forms.alert("No elements selected.", exitscript=True)
    
    dialog = ReValueDialog(elements, is_instance_mode=False)
    result = dialog.show_dialog()
    
    if result:
        revalue_map = dialog.get_revalue_map()
        is_inst_mode = dialog.is_instance_mode
        target_param_name = dialog.get_selected_original_param_name()
        
        if revalue_map:
            revalue_elements(revalue_map, is_inst_mode, target_param_name)
        else:
            forms.alert("No changes to apply.")
