# -*- coding: utf-8 -*-
"""
sort_param_utils.py
-------------------
Shared utilities for sorting family parameters alphabetically within groups.

Consumed by:
  - Family Management.panel / Sort Other Parameters.pushbutton
  - Family Management.panel / Create Edit Parameters.pushbutton  (Sort tab)

Import pattern:

    from sort_param_utils import (
        get_current_parameter_order,
        supports_reorder,
        group_parameters,
        is_other_group,
        apply_parameter_sort,
    )
"""

from Autodesk.Revit.DB import FamilyParameter, LabelUtils, Transaction
from System.Collections.Generic import List

from family_param_utils import param_name


# ---------------------------------------------------------------------------
# Group info helpers
# ---------------------------------------------------------------------------

def get_group_label_from_enum(group_enum):
    """Return a human-readable label for a BuiltInParameterGroup enum value."""
    try:
        return LabelUtils.GetLabelFor(group_enum) or ""
    except Exception:
        return ""


def get_group_info(fp):
    """Return (key, label) identifying the parameter group of *fp*.

    key  – a stable string used as a dict key (not shown to the user)
    label – the human-readable group name shown in the UI
    """
    try:
        if not fp or not fp.Definition:
            return ("unknown", "Ungrouped")

        # Legacy API: BuiltInParameterGroup on the definition (Revit ≤2022)
        group_enum = getattr(fp.Definition, "ParameterGroup", None)
        if group_enum is not None:
            label = get_group_label_from_enum(group_enum) or str(group_enum)
            key = "enum:{}".format(str(group_enum))
            return (key, label)

        # Newer API: group type id (Revit 2023+)
        get_group_type_id = getattr(fp.Definition, "GetGroupTypeId", None)
        if callable(get_group_type_id):
            group_type_id = get_group_type_id()
            if group_type_id is not None:
                group_type_id_text = str(group_type_id)
                label = group_type_id_text
                get_label_for_group = getattr(LabelUtils, "GetLabelForGroup", None)
                if callable(get_label_for_group):
                    try:
                        label = get_label_for_group(group_type_id) or group_type_id_text
                    except Exception:
                        label = group_type_id_text
                key = "gtid:{}".format(group_type_id_text)
                return (key, label)
    except Exception:
        return ("unknown", "Ungrouped")

    return ("unknown", "Ungrouped")


# ---------------------------------------------------------------------------
# FamilyManager order helpers
# ---------------------------------------------------------------------------

def supports_reorder(fm):
    """Return True if FamilyManager.ReorderParameters is available."""
    method = getattr(fm, "ReorderParameters", None)
    return callable(method)


def get_current_parameter_order(fm):
    """Return parameters in current Family Types dialog order, or None if unsupported.

    Uses GetParameters() which preserves the user's manual order, rather than
    fm.Parameters which may use ElementId-based ordering.
    """
    get_parameters = getattr(fm, "GetParameters", None)
    if callable(get_parameters):
        try:
            ordered = [p for p in get_parameters()]
            if ordered:
                return ordered
        except Exception:
            pass
    return None


# ---------------------------------------------------------------------------
# Grouping helpers
# ---------------------------------------------------------------------------

def group_parameters(all_params):
    """Return (grouped, labels).

    grouped: {key: [FamilyParameter, ...]}
    labels:  {key: human_readable_label}
    """
    grouped = {}
    labels = {}
    for p in all_params:
        key, label = get_group_info(p)
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(p)
        labels[key] = label
    return grouped, labels


def is_other_group(group_key, group_label):
    """Return True if this is the "Other" group — used to pre-tick it in UI."""
    label_text = (group_label or "").strip().lower()
    return label_text == "other"


# ---------------------------------------------------------------------------
# Sort logic
# ---------------------------------------------------------------------------

def build_reordered_list(all_params, selected_keys, sort_mode, descending):
    """Return (sorted_by_group, reordered_params).

    sorted_by_group  – {key: [FamilyParameter, ...]} for each selected group
    reordered_params – full param list with selected groups sorted in-place;
                       non-selected groups remain in their original positions
    """
    grouped, _ = group_parameters(all_params)
    sorted_by_group = {}
    group_index = {}

    for key in selected_keys:
        group_items = grouped.get(key, [])

        if sort_mode == "Type then Instance":
            type_params = [p for p in group_items if not getattr(p, "IsInstance", False)]
            instance_params = [p for p in group_items if getattr(p, "IsInstance", False)]
            type_sorted = sorted(type_params, key=lambda p: param_name(p).lower(), reverse=descending)
            instance_sorted = sorted(instance_params, key=lambda p: param_name(p).lower(), reverse=descending)
            sorted_by_group[key] = type_sorted + instance_sorted
        else:
            sorted_by_group[key] = sorted(
                group_items,
                key=lambda p: param_name(p).lower(),
                reverse=descending,
            )

        group_index[key] = 0

    reordered = []
    for fp in all_params:
        key, _ = get_group_info(fp)
        if key in selected_keys:
            idx = group_index[key]
            reordered.append(sorted_by_group[key][idx])
            group_index[key] = idx + 1
        else:
            reordered.append(fp)

    return sorted_by_group, reordered


# ---------------------------------------------------------------------------
# High-level: sort inside a transaction
# ---------------------------------------------------------------------------

def apply_parameter_sort(doc, fm, selected_keys, sort_mode, descending):
    """Sort selected parameter groups in a Revit transaction.

    Returns (success, result_code, changed_group_labels, processed_count).

    result_code values:
      "ok"             – sort applied
      "already_sorted" – no changes needed (already in order)
      any other str    – error message
    """
    all_params = get_current_parameter_order(fm)
    if all_params is None:
        return (
            False,
            "FamilyManager.GetParameters is not available in this Revit version.",
            [],
            0,
        )

    if not supports_reorder(fm):
        return (
            False,
            "FamilyManager.ReorderParameters is not available in this Revit version.",
            [],
            0,
        )

    grouped, labels = group_parameters(all_params)
    sorted_by_group, reordered_params = build_reordered_list(
        all_params, selected_keys, sort_mode, descending
    )

    changed_group_labels = []
    processed_count = 0
    for key in selected_keys:
        original_names = [param_name(p) for p in grouped.get(key, [])]
        sorted_names = [param_name(p) for p in sorted_by_group.get(key, [])]
        if original_names != sorted_names:
            changed_group_labels.append(labels.get(key, key))
        processed_count += len(grouped.get(key, []))

    if not changed_group_labels:
        return True, "already_sorted", changed_group_labels, processed_count

    ordered_list = List[FamilyParameter]()
    for p in reordered_params:
        ordered_list.Add(p)

    txn = Transaction(doc, "Sort Parameter Groups")
    try:
        txn.Start()
        fm.ReorderParameters(ordered_list)
        txn.Commit()
    except Exception as ex:
        try:
            txn.RollBack()
        except Exception:
            pass
        return False, str(ex), [], 0

    return True, "ok", changed_group_labels, processed_count
