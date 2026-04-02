# -*- coding: utf-8 -*-
"""
family_param_utils.py
---------------------
Shared utility functions for Family Parameter analysis.

Consumed by:
  - Family Management.panel / Purge Parameters.pushbutton
  - Family Management.panel / Create Edit Parameters.pushbutton
  (and any future tool in this extension that needs family-parameter helpers)

Import pattern (pyRevit adds <extension>/lib/ to sys.path automatically):

    from family_param_utils import (
        get_family_parameters,
        find_directly_used_params,
        ...
    )
"""

import re

from Autodesk.Revit.DB import (
    Dimension,
    FilteredElementCollector,
    LabelUtils,
    LinearArray,
    RadialArray,
)

# Maximum number of BFS depth levels ever rendered (mirrors PurgeParameters cap).
MAX_DEPTH_CAP = 10


# ---------------------------------------------------------------------------
# Basic accessors
# ---------------------------------------------------------------------------

def get_family_parameters(fm):
    """Return a list of all FamilyParameter objects from FamilyManager *fm*."""
    getter = getattr(fm, "GetParameters", None)
    if callable(getter):
        try:
            return list(getter())
        except Exception:
            pass
    try:
        return list(fm.Parameters)
    except Exception:
        return []


def safe_formula(fp):
    """Return the formula string for *fp*, or '' on any error."""
    try:
        return fp.Formula or ""
    except Exception:
        return ""


def param_name(fp):
    """Return the Definition.Name of a FamilyParameter, or '' on error."""
    try:
        return fp.Definition.Name or ""
    except Exception:
        return ""


def group_label(fp):
    """Return a human-readable group label for a FamilyParameter."""
    try:
        g = fp.Definition.ParameterGroup
    except Exception:
        return ""
    try:
        return LabelUtils.GetLabelFor(g)
    except Exception:
        try:
            return str(g)
        except Exception:
            return ""


def data_type_label(fp):
    """Return a human-readable data-type string for a FamilyParameter."""
    try:
        dt = fp.Definition.ParameterType
        return str(dt)
    except Exception:
        pass
    try:
        dt = fp.Definition.GetDataType()
        return LabelUtils.GetLabelForSpec(dt)
    except Exception:
        pass
    return ""


def formula_references_parameter(formula_text, parameter_name):
    """Return True if *formula_text* contains a whole-word reference to *parameter_name*."""
    if not formula_text or not parameter_name:
        return False
    try:
        pattern = r"(?<![A-Za-z0-9_]){}(?![A-Za-z0-9_])".format(re.escape(parameter_name))
        return re.search(pattern, formula_text, flags=re.IGNORECASE) is not None
    except Exception:
        return parameter_name.lower() in formula_text.lower()


def is_family_type_parameter(fp):
    """Return True if *fp* is a FamilyType parameter (should never be purged or deleted)."""
    try:
        from Autodesk.Revit.DB import ParameterType
        return fp.Definition.ParameterType == ParameterType.FamilyType
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import SpecTypeId
        return fp.Definition.GetDataType() == SpecTypeId.Reference.FamilyType
    except Exception:
        pass
    return "familytype" in data_type_label(fp).lower().replace(" ", "")


def is_system_parameter(fp):
    """Return True if *fp* is a built-in / system parameter (negative ElementId).

    System params like 'Model', 'Manufacturer', 'URL' etc. are built-in and
    should never be shown in purge/delete UIs.
    """
    try:
        pid = fp.Id
        try:
            int_val = int(pid.Value)        # Revit 2024+ (Int64)
        except AttributeError:
            int_val = int(pid.IntegerValue)
        return int_val < 0
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Core "in-use" detection
# ---------------------------------------------------------------------------

def _revit_obj_name(obj):
    """Safely get a name from a Revit object that might be a .NET null.

    In IronPython, methods returning a null Revit reference type come back as
    a Python object that is NOT None but has no usable data.  Accessing any
    real property on it (like .Id) throws, which we use to detect nulls.
    """
    if obj is None:
        return None
    try:
        _ = obj.Id      # throws on .NET null wrappers
        return param_name(obj) or None
    except Exception:
        return None


def find_directly_used_params(doc, fm):
    """Return the set of FamilyParameter objects that are directly in use.

    Uses only read-only API calls — no transaction tests.

    Pass 1 – Dimension labels: any parameter assigned as a dimension label is
             directly in use.

    Pass 2 – GetAssociatedFamilyParameter: for every non-type element in the
             document, iterates its parameters and calls
             FamilyManager.GetAssociatedFamilyParameter(param).  This covers
             nested family instance parameters, array-count parameters,
             element visibility conditions, and Yes/No parameters.

    Pass 3 – Array count & label associations on LinearArray / RadialArray.
    """
    all_params = get_family_parameters(fm)
    directly_used_names = set()

    # Pass 1: dimension labels
    try:
        dims = FilteredElementCollector(doc).OfClass(Dimension).ToElements()
        for dim in dims:
            try:
                lbl = dim.FamilyLabel
                if lbl is not None:
                    directly_used_names.add(param_name(lbl))
            except Exception:
                pass
    except Exception:
        pass

    # Pass 2: element parameter associations
    try:
        elements = (
            FilteredElementCollector(doc)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for elem in elements:
            try:
                for p in elem.Parameters:
                    try:
                        assoc_fp = fm.GetAssociatedFamilyParameter(p)
                        n = _revit_obj_name(assoc_fp)
                        if n:
                            directly_used_names.add(n)
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass

    # Pass 3: array count / label associations
    for array_class in (LinearArray, RadialArray):
        try:
            arrays = (
                FilteredElementCollector(doc)
                .OfClass(array_class)
                .WhereElementIsNotElementType()
                .ToElements()
            )
            for arr in arrays:
                for attr in ("FamilyParameterForCount", "Label"):
                    try:
                        n = _revit_obj_name(getattr(arr, attr, None))
                        if n:
                            directly_used_names.add(n)
                    except Exception:
                        pass
        except Exception:
            pass

    return set(fp for fp in all_params if param_name(fp) in directly_used_names)


def find_formula_referencing_params(fm, target_param_name):
    """Return a list of parameter names whose formulas reference *target_param_name*.

    Used by the Delete button to warn the user which other parameters will
    have broken formulas if *target_param_name* is deleted.
    """
    referencing = []
    for fp in get_family_parameters(fm):
        f = safe_formula(fp)
        if f and formula_references_parameter(f, target_param_name):
            referencing.append(param_name(fp))
    return referencing


# ---------------------------------------------------------------------------
# BFS depth analysis (used by Purge Parameters)
# ---------------------------------------------------------------------------

def build_depth_analysis(fm, directly_used, max_depth):
    """BFS from the directly-used set following formula dependencies.

    At each depth level we look for parameters (not yet marked unsafe) whose
    name appears in the formula of a parameter in the current unsafe frontier.
    Repeats up to *max_depth* times.

    Returns (safe_list, unsafe_name_set).
    """
    all_params = get_family_parameters(fm)
    name_to_formula = {param_name(fp): safe_formula(fp) for fp in all_params}

    unsafe_names = set(param_name(fp) for fp in directly_used)
    frontier = set(unsafe_names)

    for _depth in range(max_depth):
        next_frontier = set()
        for fp in all_params:
            name = param_name(fp)
            if name in unsafe_names:
                continue
            for f_name in frontier:
                formula = name_to_formula.get(f_name, "")
                if formula and formula_references_parameter(formula, name):
                    next_frontier.add(name)
                    break
        if not next_frontier:
            break
        unsafe_names |= next_frontier
        frontier = next_frontier

    safe_params = [
        fp for fp in all_params
        if param_name(fp) not in unsafe_names
        and not is_family_type_parameter(fp)
        and not is_system_parameter(fp)
    ]
    return safe_params, unsafe_names


def compute_reverse_deps(safe_params, max_depth):
    """Within the safe pool, compute reverse formula dependencies at each depth hop.

    Returns a dict: {param_name: {1: "name1, name2", 2: "name3", ...}}.

    Depth 1 = safe params whose own formula directly references this param.
    Depth 2 = safe params whose formula references a D1 intermediate, etc.
    """
    safe_names = set(param_name(fp) for fp in safe_params)
    name_to_formula = {param_name(fp): safe_formula(fp) for fp in safe_params}

    # direct_refs[a] = set of safe param names that appear in formula of 'a'
    direct_refs = {}
    for fp in safe_params:
        name = param_name(fp)
        formula = name_to_formula.get(name, "")
        direct_refs[name] = set()
        if formula:
            for other in safe_names:
                if other != name and formula_references_parameter(formula, other):
                    direct_refs[name].add(other)

    # reverse_refs[b] = set of safe param names whose formula references 'b'
    reverse_refs = {name: set() for name in safe_names}
    for a, refs in direct_refs.items():
        for b in refs:
            reverse_refs[b].add(a)

    result = {name: {} for name in safe_names}

    for pname in safe_names:
        visited = {pname}
        current_level = reverse_refs.get(pname, set()) - visited
        for depth in range(1, max_depth + 1):
            if not current_level:
                break
            result[pname][depth] = ", ".join(sorted(current_level))
            visited |= current_level
            next_level = set()
            for n in current_level:
                for parent in reverse_refs.get(n, set()):
                    if parent not in visited:
                        next_level.add(parent)
            current_level = next_level

    return result
