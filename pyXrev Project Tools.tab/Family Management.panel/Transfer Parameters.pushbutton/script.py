# -*- coding: utf-8 -*-
"""
Transfer Family Parameters
--------------------------
Reads parameters from the current (source) family and transfers selected ones
to other open family documents.

- Maintains instance/type and parameter group.
- Shared parameters are located in the active shared param file by GUID.
  If the GUID is absent, the definition is exported to a "pyEXPORT" group
  so it can then be imported into the target family.
- Parameters with formulas are highlighted in the UI and can be filtered out.
- Results are reported in the pyRevit output panel.

Supports Revit 2020 – 2026 via ForgeTypeId / ParameterType dual-path.
"""
# ---------------------------------------------------------------------------
# Imports
# ---------------------------------------------------------------------------
from Autodesk.Revit.DB import (
    BuiltInParameter,
    Category,
    ElementId,
    Family,
    FamilySymbol,
    FamilySource,
    FilteredElementCollector,
    IFamilyLoadOptions,
    Material,
    StorageType,
    SubTransaction,
    Transaction,
    LabelUtils,
    CategoryType,
    GraphicsStyleType,
    ExternalDefinitionCreationOptions,
)
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from pyrevit import revit, forms, script

# ---------------------------------------------------------------------------
# Active document / application guards
# ---------------------------------------------------------------------------
doc   = revit.doc
uiapp = __revit__                   # noqa: F821  (pyRevit injects __revit__)
app   = uiapp.Application

if not doc.IsFamilyDocument:
    forms.alert(
        "This tool must be run from the Family Editor.\n"
        "Open a family (*.rfa) for editing first.",
        exitscript=True
    )

output = script.get_output()


class _FamilyLoadOptions(IFamilyLoadOptions):
    """Always load/overwrite when loading nested families into targets."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = FamilySource.Family
        overwriteParameterValues.Value = True
        return True

# ---------------------------------------------------------------------------
# Version-safe helpers
# ---------------------------------------------------------------------------

def _get_data_type(definition):
    """Return the data type for a parameter definition.

    Revit 2022+ exposes GetDataType() → ForgeTypeId.
    Older versions use .ParameterType (enum).
    Returns whatever the current API version accepts for AddParameter calls.
    """
    try:
        return definition.GetDataType()          # ForgeTypeId  (2022+)
    except AttributeError:
        return definition.ParameterType          # ParameterType enum (2020/2021)


def _group_of(family_param):
    """Return the BuiltInParameterGroup (or ForgeTypeId group) for the param.

    Revit 2025 deprecated BuiltInParameterGroup but ParameterGroup still
    works via compatibility shims through at least 2026 preview.
    """
    return family_param.Definition.ParameterGroup


def _group_label(family_param):
    """Human-readable group name, e.g. 'Dimensions'."""
    try:
        return LabelUtils.GetLabelFor(family_param.Definition.ParameterGroup)
    except Exception:
        try:
            return str(family_param.Definition.ParameterGroup)
        except Exception:
            return ""


def _is_likely_nontransferable(family_param):
    """Heuristic marker for built-in/system parameters.

    These are often hard-coded by Revit and typically should not be re-added,
    but formulas can still be copied onto existing counterparts in targets.
    """
    try:
        bip = family_param.Definition.BuiltInParameter
        if bip != BuiltInParameter.INVALID:
            return True
    except Exception:
        pass
    return False


def _is_family_type_parameter(family_param):
    """Best-effort detection for Family Type parameters."""
    # Legacy API path
    try:
        ptype = family_param.Definition.ParameterType
        if str(ptype) == "FamilyType":
            return True
    except Exception:
        pass

    # New API path
    try:
        data_type = family_param.Definition.GetDataType()
        if data_type:
            type_id = getattr(data_type, "TypeId", "")
            if type_id and "autodesk.revit.category" in type_id.lower():
                return True
    except Exception:
        pass

    return False


def _try_get_familytype_category(target_doc, family_param):
    """Try to resolve the family type category for category-based AddParameter."""
    try:
        data_type = family_param.Definition.GetDataType()
    except Exception:
        data_type = None

    if data_type is None:
        return None

    # Revit 2022+ supports Category.GetCategory(doc, ForgeTypeId)
    try:
        cat = Category.GetCategory(target_doc, data_type)
        if cat:
            return cat
    except Exception:
        pass

    # Fallback parse if TypeId encodes OST_ name (best-effort)
    try:
        type_id = data_type.TypeId
        if type_id and "ost_" in type_id.lower():
            token = type_id.split(".")[-1]
            import System
            bic_type = System.Type.GetType("Autodesk.Revit.DB.BuiltInCategory, RevitAPI")
            if bic_type:
                bic = System.Enum.Parse(bic_type, token, True)
                return Category.GetCategory(target_doc, bic)
    except Exception:
        pass

    return None


def _category_by_integer_id(target_doc, category_int_id):
    """Return Category in target_doc matching category_int_id, else None."""
    if category_int_id is None:
        return None
    try:
        for cat in target_doc.Settings.Categories:
            if cat.Id.IntegerValue == int(category_int_id):
                return cat
    except Exception:
        pass
    return None


def _detect_familytype_category_id(source_doc, source_fm, family_param):
    """Infer Family Type parameter category by inspecting type-assigned values.

    For Family Type params, FamilyType.AsElementId(family_param) can point to a
    FamilySymbol/ElementType whose Category gives the required AddParameter
    overload category.
    """
    try:
        for fam_type in source_fm.Types:
            try:
                val_id = fam_type.AsElementId(family_param)
            except Exception:
                continue
            if val_id is None or val_id == ElementId.InvalidElementId:
                continue

            elem = source_doc.GetElement(val_id)
            if elem is not None and elem.Category is not None:
                return elem.Category.Id.IntegerValue
    except Exception:
        pass
    return None


def _target_has_types_for_category(target_doc, category, cache_dict):
    """Return True if target_doc has at least one element type in category."""
    if category is None:
        return False
    cat_id = category.Id.IntegerValue
    if cat_id in cache_dict:
        return cache_dict[cat_id]

    has_any = False
    try:
        types = (
            FilteredElementCollector(target_doc)
            .WhereElementIsElementType()
            .ToElements()
        )
        for t in types:
            if t.Category is not None and t.Category.Id.IntegerValue == cat_id:
                has_any = True
                break
    except Exception:
        has_any = False

    cache_dict[cat_id] = has_any
    return has_any


def _collect_source_nested_families_for_param(source_doc, source_fm, family_param, expected_category_id=None):
    """Collect nested Family elements referenced by a Family Type parameter values."""
    fam_map = {}

    try:
        for fam_type in source_fm.Types:
            try:
                val_id = fam_type.AsElementId(family_param)
            except Exception:
                continue

            if val_id is None or val_id == ElementId.InvalidElementId:
                continue

            elem = source_doc.GetElement(val_id)
            if elem is None:
                continue

            cat = getattr(elem, "Category", None)
            if expected_category_id is not None:
                if cat is None or cat.Id.IntegerValue != int(expected_category_id):
                    continue

            fam = getattr(elem, "Family", None)
            if fam is None:
                continue

            fam_map[fam.Id.IntegerValue] = fam
    except Exception:
        pass

    return list(fam_map.values())


def _preload_familytype_sources_into_target(source_doc, source_fm, target_doc, param_item,
                                            category_type_cache, results):
    """Load source nested families into target to satisfy Family Type parameters."""
    fp = param_item.FamilyParameter
    name = param_item.Name

    cat = _category_by_integer_id(target_doc, param_item.FamilyTypeCategoryId)
    if cat is None:
        cat = _try_get_familytype_category(target_doc, fp)
    if cat is None:
        return

    if _target_has_types_for_category(target_doc, cat, category_type_cache):
        return

    source_nested_families = _collect_source_nested_families_for_param(
        source_doc,
        source_fm,
        fp,
        expected_category_id=param_item.FamilyTypeCategoryId,
    )

    if not source_nested_families:
        results["preload_notes"].append(
            "{}: no source nested families found to preload for category '{}'".format(name, cat.Name)
        )
        return

    load_opts = _FamilyLoadOptions()

    loaded_count = 0
    for nested_family in source_nested_families:
        nested_doc = None
        try:
            nested_doc = source_doc.EditFamily(nested_family)
            nested_doc.LoadFamily(target_doc, load_opts)
            loaded_count += 1
        except Exception as ex:
            results["preload_notes"].append(
                "{}: could not preload nested family '{}': {}".format(
                    name,
                    nested_family.Name,
                    ex,
                )
            )
        finally:
            if nested_doc:
                try:
                    nested_doc.Close(False)
                except Exception:
                    pass

    # Reset cache entry, then re-evaluate.
    category_type_cache.pop(cat.Id.IntegerValue, None)

    if loaded_count > 0:
        results["preload_notes"].append(
            "{}: preloaded {} nested family/families for category '{}'".format(
                name,
                loaded_count,
                cat.Name,
            )
        )


# ---------------------------------------------------------------------------
# Data model: ParamItem (INotifyPropertyChanged for WPF binding)
# ---------------------------------------------------------------------------

class ParamItem(INotifyPropertyChanged):
    """Represents one source family parameter shown in the WPF ListView."""

    def __init__(self, fp):
        self._is_selected   = True
        self._fp            = fp           # live FamilyParameter reference
        self.Name           = fp.Definition.Name
        self.GroupLabel     = _group_label(fp)
        self.InstanceTypeLabel = "Instance" if fp.IsInstance else "Type"
        self.Formula        = fp.Formula or ""
        self.HasFormula     = bool(self.Formula)
        self.IsShared       = fp.IsShared
        self.SharedLabel    = "Yes" if fp.IsShared else ""
        self.IsLikelyNonTransferable = _is_likely_nontransferable(fp)
        try:
            self.Guid       = fp.GUID      # only valid for shared params
        except Exception:
            self.Guid       = None
        self.FamilyTypeCategoryId = None
        self._changed_handlers = []

    # INotifyPropertyChanged
    def add_PropertyChanged(self, handler):
        self._changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._changed_handlers:
            self._changed_handlers.remove(handler)

    def _notify(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for h in self._changed_handlers:
            h(self, args)

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify("IsSelected")

    @property
    def FamilyParameter(self):
        return self._fp


class TargetItem(INotifyPropertyChanged):
    """Represents one target family document shown in the WPF ListView."""

    def __init__(self, revit_doc):
        self._is_selected = True
        self._doc         = revit_doc
        self.Title        = revit_doc.Title
        self._changed_handlers = []

    def add_PropertyChanged(self, handler):
        self._changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._changed_handlers:
            self._changed_handlers.remove(handler)

    def _notify(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for h in self._changed_handlers:
            h(self, args)

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify("IsSelected")

    @property
    def Document(self):
        return self._doc


class NestedFamilyItem(INotifyPropertyChanged):
    """Selectable nested family candidate in Nested Families tab."""

    def __init__(self, family_elem):
        self._is_selected = False
        self._family = family_elem
        self.Name = family_elem.Name
        cat = getattr(family_elem, "FamilyCategory", None)
        self.CategoryName = cat.Name if cat is not None else "<No Category>"
        self.Status = "Available"
        self._changed_handlers = []

    def add_PropertyChanged(self, handler):
        self._changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._changed_handlers:
            self._changed_handlers.remove(handler)

    def _notify(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for h in self._changed_handlers:
            h(self, args)

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify("IsSelected")

    @property
    def Family(self):
        return self._family


class ObjectStyleItem(INotifyPropertyChanged):
    """Selectable object style candidate in Object Styles tab."""

    def __init__(self, parent_category, subcategory_name, style_type, source_subcategory):
        self._is_selected = False
        self.ParentCategory = parent_category
        self.Name = subcategory_name
        self.StyleType = style_type
        self.SourceSubcategory = source_subcategory
        self.Status = "Available"
        self._changed_handlers = []

    def add_PropertyChanged(self, handler):
        self._changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._changed_handlers:
            self._changed_handlers.remove(handler)

    def _notify(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for h in self._changed_handlers:
            h(self, args)

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify("IsSelected")


# ---------------------------------------------------------------------------
# Shared parameter file helpers
# ---------------------------------------------------------------------------

def _open_def_file():
    """Return the active DefinitionFile, or None if none is configured."""
    try:
        return app.OpenSharedParameterFile()
    except Exception:
        return None


def _find_definition_by_guid(def_file, guid):
    """Search all groups in def_file for a definition whose GUID matches."""
    if def_file is None or guid is None:
        return None
    try:
        for grp in def_file.Groups:
            for defn in grp.Definitions:
                try:
                    if defn.GUID == guid:
                        return defn
                except Exception:
                    pass
    except Exception:
        pass
    return None


def ensure_in_sharedparam_file(param_item):
    """Ensure the shared parameter definition exists in the active .txt file.

    If the GUID is already present, return the ExternalDefinition.
    If not, create it in a 'pyEXPORT' group and return the new definition.
    Returns None if there is no shared parameter file configured.

    Args:
        param_item (ParamItem): the source parameter item (must be shared).

    Returns:
        ExternalDefinition or None
    """
    def_file = _open_def_file()
    if def_file is None:
        return None

    # Try to find by GUID
    existing = _find_definition_by_guid(def_file, param_item.Guid)
    if existing:
        return existing

    # Not found — export to pyEXPORT group
    fp = param_item.FamilyParameter
    data_type = _get_data_type(fp.Definition)

    # Get or create the pyEXPORT group
    export_group = def_file.Groups["pyEXPORT"]
    if export_group is None:
        export_group = def_file.Groups.Create("pyEXPORT")

    opts = ExternalDefinitionCreationOptions(param_item.Name, data_type)
    opts.GUID            = param_item.Guid
    opts.Visible         = True
    opts.UserModifiable  = True
    try:
        opts.Description = fp.Definition.Description
    except Exception:
        pass

    new_defn = export_group.Definitions.Create(opts)
    return new_defn


# ---------------------------------------------------------------------------
# Parameter transfer
# ---------------------------------------------------------------------------

def _existing_param_names(family_manager):
    """Return a set of parameter names already in a FamilyManager."""
    names = set()
    for fp in family_manager.Parameters:
        names.add(fp.Definition.Name)
    return names


def _find_param_by_name(family_manager, name):
    """Return FamilyParameter by definition name, else None."""
    for fp in family_manager.Parameters:
        if fp.Definition.Name == name:
            return fp
    return None


def _safe_elem_name(elem):
    """Best-effort element type name extraction across API oddities."""
    if elem is None:
        return None

    try:
        name = elem.Name
        if name:
            return name
    except Exception:
        pass

    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            name = p.AsString()
            if name:
                return name
    except Exception:
        pass

    try:
        p = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p:
            name = p.AsString()
            if name:
                return name
    except Exception:
        pass

    return None


def _safe_elem_family_name(elem):
    """Best-effort family name extraction for element types."""
    if elem is None:
        return None

    try:
        fam_name = elem.FamilyName
        if fam_name:
            return fam_name
    except Exception:
        pass

    try:
        fam = getattr(elem, "Family", None)
        if fam is not None and fam.Name:
            return fam.Name
    except Exception:
        pass

    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if p:
            fam_name = p.AsString()
            if fam_name:
                return fam_name
    except Exception:
        pass

    return None


def _safe_category_name(cat):
    """Best-effort category name extraction for diagnostics."""
    if cat is None:
        return "<No Category>"
    try:
        if cat.Name:
            return cat.Name
    except Exception:
        pass
    try:
        return "CategoryId {}".format(cat.Id.IntegerValue)
    except Exception:
        return "<Unknown Category>"


def _get_or_build_type_cache(target_doc, cache_dict):
    """Cache target element types keyed by category/name and family/type."""
    key = "_types_by_cat_and_name"
    key_family = "_types_by_cat_family_and_name"
    if key in cache_dict and key_family in cache_dict:
        return cache_dict[key], cache_dict[key_family]

    if key in cache_dict:
        return cache_dict[key], cache_dict.get(key_family, {})

    mapping = {}
    mapping_family = {}
    try:
        for elem_type in (
            FilteredElementCollector(target_doc)
            .WhereElementIsElementType()
            .ToElements()
        ):
            cat = getattr(elem_type, "Category", None)
            if cat is None:
                continue
            type_name = _safe_elem_name(elem_type)
            if not type_name:
                continue

            k = (cat.Id.IntegerValue, type_name)
            if k not in mapping:
                mapping[k] = elem_type

            fam_name = _safe_elem_family_name(elem_type)
            if fam_name:
                kf = (cat.Id.IntegerValue, fam_name, type_name)
                if kf not in mapping_family:
                    mapping_family[kf] = elem_type
    except Exception:
        pass

    cache_dict[key] = mapping
    cache_dict[key_family] = mapping_family
    return mapping, mapping_family


def _get_or_build_symbol_cache(target_doc, cache_dict):
    """Cache target FamilySymbol types keyed by category/name and family/type."""
    key = "_symbols_by_cat_and_name"
    key_family = "_symbols_by_cat_family_and_name"
    if key in cache_dict and key_family in cache_dict:
        return cache_dict[key], cache_dict[key_family]

    mapping = {}
    mapping_family = {}
    try:
        for symbol in FilteredElementCollector(target_doc).OfClass(FamilySymbol).ToElements():
            cat = getattr(symbol, "Category", None)
            if cat is None:
                continue

            type_name = _safe_elem_name(symbol)
            if not type_name:
                continue

            k = (cat.Id.IntegerValue, type_name)
            if k not in mapping:
                mapping[k] = symbol

            fam_name = _safe_elem_family_name(symbol)
            if fam_name:
                kf = (cat.Id.IntegerValue, fam_name, type_name)
                if kf not in mapping_family:
                    mapping_family[kf] = symbol
    except Exception:
        pass

    cache_dict[key] = mapping
    cache_dict[key_family] = mapping_family
    return mapping, mapping_family


def _get_or_build_material_cache(target_doc, cache_dict):
    """Cache target materials by exact name."""
    key = "_materials_by_name"
    if key in cache_dict:
        return cache_dict[key]

    mapping = {}
    try:
        for mat in FilteredElementCollector(target_doc).OfClass(Material).ToElements():
            mapping[mat.Name] = mat
    except Exception:
        pass

    cache_dict[key] = mapping
    return mapping


def _get_or_build_created_materials(cache_dict):
    """Track created material names during current target transfer run."""
    key = "_created_material_names"
    if key not in cache_dict:
        cache_dict[key] = set()
    return cache_dict[key]


def _copy_material_properties(source_mat, target_mat):
    """Best-effort copy of material graphics/appearance properties."""
    if source_mat is None or target_mat is None:
        return

    # Basic graphics/identity properties commonly available in family docs.
    props = [
        "Color",
        "Transparency",
        "Shininess",
        "Smoothness",
        "UseRenderAppearanceForShading",
        "CutForegroundPatternColor",
        "CutForegroundPatternId",
        "CutBackgroundPatternColor",
        "CutBackgroundPatternId",
        "SurfaceForegroundPatternColor",
        "SurfaceForegroundPatternId",
        "SurfaceBackgroundPatternColor",
        "SurfaceBackgroundPatternId",
        "AppearanceAssetId",
    ]

    for prop in props:
        try:
            if hasattr(source_mat, prop) and hasattr(target_mat, prop):
                setattr(target_mat, prop, getattr(source_mat, prop))
        except Exception:
            # Some properties can be read-only or invalid in certain contexts.
            pass


def _ensure_material_in_target_by_name(target_doc, material_name, cache_dict, allow_create,
                                       source_material=None):
    """Get target material by name; optionally create if missing."""
    mats = _get_or_build_material_cache(target_doc, cache_dict)
    target_mat = mats.get(material_name)
    if target_mat:
        return True, target_mat, ""

    if not allow_create:
        return False, None, "Material '{}' not found in target".format(material_name)

    try:
        new_mat_id = Material.Create(target_doc, material_name)
        new_mat = target_doc.GetElement(new_mat_id)
        if new_mat is None:
            return False, None, "Material '{}' could not be created".format(material_name)

        _copy_material_properties(source_material, new_mat)

        mats[material_name] = new_mat
        _get_or_build_created_materials(cache_dict).add(material_name)
        return True, new_mat, ""
    except Exception as ex:
        return False, None, "Material '{}' creation failed: {}".format(material_name, ex)


def _map_element_id_value(source_doc, target_doc, source_elem_id, cache_dict,
                          preferred_type_name=None, require_family_symbol=False,
                          create_missing_materials=False):
    """Map source ElementId value to target ElementId where possible."""
    if source_elem_id is None or source_elem_id == ElementId.InvalidElementId:
        return True, ElementId.InvalidElementId, ""

    source_elem = source_doc.GetElement(source_elem_id)
    if source_elem is None:
        return False, None, "Source ElementId value does not resolve to an element"

    if isinstance(source_elem, Material):
        ok, target_mat, reason = _ensure_material_in_target_by_name(
            target_doc,
            source_elem.Name,
            cache_dict,
            create_missing_materials,
            source_elem,
        )
        if ok:
            return True, target_mat.Id, ""
        return False, None, reason

    cat = getattr(source_elem, "Category", None)
    if cat is not None:
        source_type_name = _safe_elem_name(source_elem)
        if not source_type_name and preferred_type_name:
            source_type_name = preferred_type_name
        if not source_type_name:
            return False, None, "Could not read source type name for ElementId mapping"

        source_family_name = _safe_elem_family_name(source_elem)
        if require_family_symbol:
            types_map, types_family_map = _get_or_build_symbol_cache(target_doc, cache_dict)
        else:
            types_map, types_family_map = _get_or_build_type_cache(target_doc, cache_dict)

        target_type = None
        if source_family_name:
            target_type = types_family_map.get((cat.Id.IntegerValue, source_family_name, source_type_name))
        if target_type is None and source_family_name and preferred_type_name and preferred_type_name != source_type_name:
            target_type = types_family_map.get((cat.Id.IntegerValue, source_family_name, preferred_type_name))
        if target_type is None:
            target_type = types_map.get((cat.Id.IntegerValue, source_type_name))
        if target_type is None and preferred_type_name and preferred_type_name != source_type_name:
            target_type = types_map.get((cat.Id.IntegerValue, preferred_type_name))
        if target_type:
            return True, target_type.Id, ""

        symbol_hint = " (FamilySymbol only)" if require_family_symbol else ""
        return False, None, (
            "Could not map type value '{}' (family '{}') in category '{}' by name{}".format(
                source_type_name,
                source_family_name or "?",
                _safe_category_name(cat),
                symbol_hint,
            )
        )

    return False, None, "Unsupported ElementId value mapping"


def _copy_value_from_source_to_target(source_doc, source_fm, source_fp,
                                      target_doc, target_fm, target_fp, cache_dict,
                                      create_missing_materials=False):
    """Copy value from source family current type to target family current type.

    Returns tuple(bool success, str reason_if_not_success).
    """
    if source_fp.Formula:
        return False, "Skipped: source parameter has formula"

    source_type = source_fm.CurrentType
    target_type = target_fm.CurrentType

    if source_type is None:
        return False, "Source family has no current type"
    if target_type is None:
        return False, "Target family has no current type"

    storage_type = source_fp.StorageType

    try:
        if storage_type == StorageType.Double:
            target_fm.Set(target_fp, source_type.AsDouble(source_fp))
        elif storage_type == StorageType.Integer:
            target_fm.Set(target_fp, source_type.AsInteger(source_fp))
        elif storage_type == StorageType.String:
            val = source_type.AsString(source_fp)
            target_fm.Set(target_fp, val if val is not None else "")
        elif storage_type == StorageType.ElementId:
            preferred_type_name = None
            try:
                preferred_type_name = source_type.AsValueString(source_fp)
            except Exception:
                preferred_type_name = None

            require_family_symbol = False
            try:
                require_family_symbol = _is_family_type_parameter(target_fp)
            except Exception:
                require_family_symbol = False

            ok, mapped_id, reason = _map_element_id_value(
                source_doc,
                target_doc,
                source_type.AsElementId(source_fp),
                cache_dict,
                preferred_type_name,
                require_family_symbol,
                create_missing_materials,
            )
            if not ok:
                return False, reason

            if require_family_symbol:
                mapped_elem = target_doc.GetElement(mapped_id)
                if not isinstance(mapped_elem, FamilySymbol):
                    return False, "Mapped value is not a FamilySymbol for a Family Type parameter"

            target_fm.Set(target_fp, mapped_id)
        else:
            return False, "Unsupported storage type: {}".format(storage_type)
    except Exception as ex:
        return False, str(ex)

    return True, ""


def transfer_params_to_doc(param_items, target_doc, transfer_formulas=True,
                           transfer_values=True, overwrite_existing_values=False,
                           create_missing_materials=False):
    """Transfer the given ParamItems into target_doc's FamilyManager.

    Returns:
        dict with keys 'transferred', 'skipped', 'failed' — each a list of
        (param_name, reason) tuples.
    """
    results = {
        "transferred": [],
        "skipped": [],
        "failed": [],
        "formula_applied": [],
        "formula_failed": [],
        "preload_notes": [],
        "values_copied": [],
        "values_skipped": [],
        "values_failed": [],
        "materials_created": [],
    }
    fm = target_doc.FamilyManager
    source_fm = doc.FamilyManager

    existing = _existing_param_names(fm)
    formula_candidates = []
    value_candidates = []
    category_type_cache = {}
    value_map_cache = {}

    t = Transaction(target_doc, "Transfer Family Parameters")
    try:
        # Preload needed nested families for Family Type parameters before starting transaction.
        for item in param_items:
            if item.Name in existing:
                continue
            if _is_family_type_parameter(item.FamilyParameter):
                _preload_familytype_sources_into_target(
                    doc,
                    doc.FamilyManager,
                    target_doc,
                    item,
                    category_type_cache,
                    results,
                )

        t.Start()
        for item in param_items:
            name = item.Name
            if name in existing:
                results["skipped"].append((name, "Already exists in target"))
                if transfer_formulas and item.Formula:
                    formula_candidates.append((item, "existing"))
                if transfer_values:
                    if overwrite_existing_values:
                        value_candidates.append((item, "existing"))
                    else:
                        results["values_skipped"].append((
                            name,
                            "Existing parameter and overwrite is disabled",
                        ))
                continue

            try:
                st = SubTransaction(target_doc)
                st.Start()

                if item.IsShared:
                    ext_defn = ensure_in_sharedparam_file(item)
                    if ext_defn is None:
                        results["failed"].append(
                            (name, "No shared parameter file configured in target session")
                        )
                        st.RollBack()
                        continue
                    fm.AddParameter(ext_defn, _group_of(item.FamilyParameter),
                                    item.FamilyParameter.IsInstance)
                else:
                    fp        = item.FamilyParameter
                    grp       = _group_of(fp)
                    is_inst   = fp.IsInstance
                    data_type = _get_data_type(fp.Definition)
                    try:
                        fm.AddParameter(name, grp, data_type, is_inst)
                    except Exception as add_ex:
                        # Family Type parameters need the category-based overload.
                        if _is_family_type_parameter(fp):
                            cat = _category_by_integer_id(target_doc, item.FamilyTypeCategoryId)
                            if cat is None:
                                cat = _try_get_familytype_category(target_doc, fp)
                            if cat is None:
                                raise Exception(
                                    "Family Type parameter requires category overload; "
                                    "could not resolve category from source values/definition"
                                )

                            # Family Type params typically require at least one loaded
                            # element type of the chosen category in the target family.
                            if not _target_has_types_for_category(target_doc, cat, category_type_cache):
                                raise Exception(
                                    "Target family has no loaded type elements in category '{}' "
                                    "required by this Family Type parameter".format(cat.Name)
                                )

                            fm.AddParameter(name, grp, cat, is_inst)
                        else:
                            raise add_ex

                # Force regen while inside the sub-transaction so failures can roll back cleanly.
                target_doc.Regenerate()
                st.Commit()

                results["transferred"].append((name, ""))
                existing.add(name)   # prevent duplicate if same name appears twice
                if transfer_formulas and item.Formula:
                    formula_candidates.append((item, "transferred"))
                if transfer_values:
                    value_candidates.append((item, "transferred"))

            except Exception as ex:
                try:
                    st.RollBack()
                except Exception:
                    pass
                results["failed"].append((name, str(ex)))

        t.Commit()

        # Value copy runs after parameter creation commit.
        if transfer_values and value_candidates:
            tv = Transaction(target_doc, "Transfer Parameter Values")
            try:
                tv.Start()
                for item, source_state in value_candidates:
                    target_fp = _find_param_by_name(fm, item.Name)
                    if target_fp is None:
                        results["values_failed"].append((
                            item.Name,
                            "Could not find matching parameter in target",
                        ))
                        continue

                    ok, reason = _copy_value_from_source_to_target(
                        doc,
                        source_fm,
                        item.FamilyParameter,
                        target_doc,
                        fm,
                        target_fp,
                        value_map_cache,
                        create_missing_materials,
                    )
                    if ok:
                        results["values_copied"].append((item.Name, source_state))
                    else:
                        if reason.startswith("Skipped:"):
                            results["values_skipped"].append((item.Name, reason.replace("Skipped: ", "")))
                        else:
                            results["values_failed"].append((item.Name, reason))
                tv.Commit()
                created = sorted(list(_get_or_build_created_materials(value_map_cache)))
                if created:
                    results["materials_created"] = created
            except Exception as value_tx_ex:
                try:
                    tv.RollBack()
                except Exception:
                    pass
                copied_names = {n for n, _ in results["values_copied"]}
                failed_names = {n for n, _ in results["values_failed"]}
                skipped_names = {n for n, _ in results["values_skipped"]}
                for item, _ in value_candidates:
                    if item.Name not in copied_names and item.Name not in failed_names and item.Name not in skipped_names:
                        results["values_failed"].append((
                            item.Name,
                            "Value transaction rolled back: " + str(value_tx_ex),
                        ))

        # Formulas must be assigned in a separate transaction after params exist.
        if transfer_formulas and formula_candidates:
            tf = Transaction(target_doc, "Transfer Parameter Formulas")
            try:
                tf.Start()
                for item, source_state in formula_candidates:
                    target_fp = None
                    for existing_fp in fm.Parameters:
                        if existing_fp.Definition.Name == item.Name:
                            target_fp = existing_fp
                            break
                    if target_fp is None:
                        results["formula_failed"].append(
                            (item.Name, "Could not find matching parameter in target")
                        )
                        continue
                    try:
                        fm.SetFormula(target_fp, item.Formula)
                        results["formula_applied"].append((item.Name, source_state))
                    except Exception as formula_ex:
                        results["formula_failed"].append((item.Name, str(formula_ex)))
                tf.Commit()
            except Exception as formula_tx_ex:
                try:
                    tf.RollBack()
                except Exception:
                    pass
                for item, _ in formula_candidates:
                    already_recorded = {n for n, _ in results["formula_applied"]}
                    failed_recorded = {n for n, _ in results["formula_failed"]}
                    if item.Name not in already_recorded and item.Name not in failed_recorded:
                        results["formula_failed"].append(
                            (item.Name, "Formula transaction rolled back: " + str(formula_tx_ex))
                        )

    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        # Mark everything not yet processed as failed
        transferred_names = {n for n, _ in results["transferred"]}
        for item in param_items:
            if item.Name not in transferred_names and \
               item.Name not in {n for n, _ in results["skipped"]} and \
               item.Name not in {n for n, _ in results["failed"]}:
                results["failed"].append((item.Name, "Transaction rolled back: " + str(ex)))

    return results


def transfer_nested_families_to_doc(source_doc, nested_items, target_doc,
                                    load_missing_only=True, reload_existing=False):
    """Transfer selected nested families from source_doc to target_doc."""
    results = {
        "loaded": [],
        "reloaded": [],
        "skipped": [],
        "failed": [],
    }

    target_family_names = set()
    try:
        for fam in FilteredElementCollector(target_doc).OfClass(Family).ToElements():
            target_family_names.add(fam.Name)
    except Exception:
        pass

    load_opts = _FamilyLoadOptions()

    for item in nested_items:
        fam = item.Family
        fam_name = item.Name
        exists_in_target = fam_name in target_family_names

        if exists_in_target and load_missing_only and not reload_existing:
            results["skipped"].append((fam_name, "Already exists in target"))
            continue

        nested_doc = None
        try:
            nested_doc = source_doc.EditFamily(fam)
            nested_doc.LoadFamily(target_doc, load_opts)

            if exists_in_target:
                results["reloaded"].append((fam_name, ""))
            else:
                results["loaded"].append((fam_name, ""))
                target_family_names.add(fam_name)
        except Exception as ex:
            results["failed"].append((fam_name, str(ex)))
        finally:
            if nested_doc:
                try:
                    nested_doc.Close(False)
                except Exception:
                    pass

    return results


def _find_parent_category_by_name(target_doc, parent_name):
    for cat in target_doc.Settings.Categories:
        if cat.Name == parent_name:
            return cat
    return None


def _find_subcategory_by_name(parent_cat, subcat_name):
    try:
        for subcat in parent_cat.SubCategories:
            if subcat.Name == subcat_name:
                return subcat
    except Exception:
        pass
    return None


def _copy_subcategory_graphics(source_subcat, target_subcat):
    # Keep best-effort to avoid hard API failures on protected slots.
    try:
        target_subcat.LineColor = source_subcat.LineColor
    except Exception:
        pass

    try:
        target_subcat.Material = source_subcat.Material
    except Exception:
        pass

    for gst in [GraphicsStyleType.Projection, GraphicsStyleType.Cut]:
        try:
            lw = source_subcat.GetLineWeight(gst)
            if lw and lw > 0:
                target_subcat.SetLineWeight(lw, gst)
        except Exception:
            pass

        try:
            lp = source_subcat.GetLinePatternId(gst)
            if lp:
                target_subcat.SetLinePatternId(lp, gst)
        except Exception:
            pass


def transfer_object_styles_to_doc(style_items, target_doc):
    """Transfer selected object styles from source family to target family."""
    results = {
        "copied": [],
        "created": [],
        "skipped": [],
        "failed": [],
    }

    tx = Transaction(target_doc, "Transfer Object Styles")
    try:
        tx.Start()
        for item in style_items:
            name = "{} : {}".format(item.ParentCategory, item.Name)

            source_subcat = item.SourceSubcategory
            if source_subcat is None:
                results["failed"].append((name, "Source subcategory unavailable"))
                continue

            parent_cat = _find_parent_category_by_name(target_doc, item.ParentCategory)
            if parent_cat is None:
                results["skipped"].append((name, "Parent category not found in target"))
                continue

            target_subcat = _find_subcategory_by_name(parent_cat, item.Name)
            if target_subcat is None:
                # System/hard-coded subcategories should not be created.
                try:
                    is_system_subcat = item.SourceSubcategory.Id.IntegerValue < 0
                except Exception:
                    is_system_subcat = False

                if is_system_subcat:
                    results["skipped"].append((
                        name,
                        "System subcategory missing in target; style copy only (no create)",
                    ))
                    continue

                try:
                    target_subcat = target_doc.Settings.Categories.NewSubcategory(parent_cat, item.Name)
                    results["created"].append((name, ""))
                except Exception as ex:
                    results["failed"].append((name, "Could not create subcategory: {}".format(ex)))
                    continue

            try:
                _copy_subcategory_graphics(source_subcat, target_subcat)
                results["copied"].append((name, ""))
            except Exception as ex:
                results["failed"].append((name, str(ex)))

        tx.Commit()
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        for item in style_items:
            name = "{} : {}".format(item.ParentCategory, item.Name)
            already = {n for n, _ in results["copied"]} | {n for n, _ in results["created"]} | {n for n, _ in results["skipped"]} | {n for n, _ in results["failed"]}
            if name not in already:
                results["failed"].append((name, "Transaction rolled back: " + str(ex)))

    return results


# ---------------------------------------------------------------------------
# WPF Dialog
# ---------------------------------------------------------------------------

class TransferParamsWindow(forms.WPFWindow):
    """Main dialog for selecting parameters and target families."""

    def __init__(self, param_items, target_items, nested_items, style_items):
        forms.WPFWindow.__init__(self, "TransferParameters.xaml")
        self._all_params   = list(param_items)
        self._target_items = list(target_items)
        self._all_nested_items = list(nested_items)
        self._all_style_items = list(style_items)
        self._all_style_parent_categories = sorted(
            {s.ParentCategory for s in self._all_style_items if s.ParentCategory}
        )

        # Populate params list
        self._params_collection = ObservableCollection[object]()
        for p in self._all_params:
            self._params_collection.Add(p)
        self.paramListView.ItemsSource = self._params_collection

        # Populate targets list
        self._targets_collection = ObservableCollection[object]()
        for t in self._target_items:
            self._targets_collection.Add(t)
        self.targetListView.ItemsSource = self._targets_collection

        # Populate nested families list
        self._nested_collection = ObservableCollection[object]()
        for n in self._all_nested_items:
            self._nested_collection.Add(n)
        self.nestedFamilyListView.ItemsSource = self._nested_collection

        # Populate object styles list
        self._styles_collection = ObservableCollection[object]()
        for s in self._all_style_items:
            self._styles_collection.Add(s)
        self.objectStylesListView.ItemsSource = self._styles_collection

        # Parent category filter for Object Styles
        self.parentStyleCategoryCombo.Items.Add("<All>")
        for pcat in self._all_style_parent_categories:
            self.parentStyleCategoryCombo.Items.Add(pcat)
        self.parentStyleCategoryCombo.SelectedIndex = 0

        # Result properties (set when OK is clicked)
        self.selected_params  = []
        self.selected_targets = []
        self.selected_nested_families = []
        self.selected_object_styles = []
        self.transfer_formulas = True
        self.transfer_values = True
        self.overwrite_existing_values = False
        self.create_missing_materials = False
        self.nested_load_missing_only = True
        self.nested_reload_existing = False
        self._apply_param_filters()
        self._apply_style_filters()

    # ------------------------------------------------------------------
    # Filter: hide-formulas toggle
    # ------------------------------------------------------------------

    def _apply_param_filters(self):
        hide = bool(self.chkHideFormulas.IsChecked)
        search = (self.txtSearch.Text or "").strip().lower()

        self._params_collection.Clear()
        for p in self._all_params:
            if hide and p.HasFormula:
                # Hidden formula params are auto-unselected to prevent transfer.
                p.IsSelected = False
                continue

            if search:
                name_match = search in (p.Name or "").lower()
                formula_match = search in (p.Formula or "").lower()
                if not name_match and not formula_match:
                    continue

            self._params_collection.Add(p)

    def on_filter_changed(self, sender, args):
        self._apply_param_filters()

    def on_search_changed(self, sender, args):
        self._apply_param_filters()

    # ------------------------------------------------------------------
    # Nested Families select all / none
    # ------------------------------------------------------------------

    def on_nested_select_all(self, sender, args):
        for n in self._nested_collection:
            n.IsSelected = True

    def on_nested_select_none(self, sender, args):
        for n in self._nested_collection:
            n.IsSelected = False

    # ------------------------------------------------------------------
    # Object Styles filtering and select all / none
    # ------------------------------------------------------------------

    def _apply_style_filters(self):
        include_model = bool(self.chkStylesIncludeModel.IsChecked)
        include_anno = bool(self.chkStylesIncludeAnnotation.IsChecked)
        include_imports = bool(self.chkStylesIncludeImports.IsChecked)
        selected_parent = self.parentStyleCategoryCombo.SelectedItem
        selected_parent = str(selected_parent) if selected_parent else "<All>"

        self._styles_collection.Clear()
        for s in self._all_style_items:
            if selected_parent != "<All>" and s.ParentCategory != selected_parent:
                continue

            t = (s.StyleType or "").lower()
            if t == "model" and not include_model:
                continue
            if t == "annotation" and not include_anno:
                continue
            if t == "import" and not include_imports:
                continue
            self._styles_collection.Add(s)

    def on_style_filter_changed(self, sender, args):
        self._apply_style_filters()

    def on_parent_style_category_changed(self, sender, args):
        self._apply_style_filters()

    def on_styles_select_all(self, sender, args):
        for s in self._styles_collection:
            s.IsSelected = True

    def on_styles_select_none(self, sender, args):
        for s in self._styles_collection:
            s.IsSelected = False

    # ------------------------------------------------------------------
    # Param select all / none
    # ------------------------------------------------------------------

    def on_params_select_all(self, sender, args):
        for p in self._params_collection:
            p.IsSelected = True

    def on_params_select_none(self, sender, args):
        for p in self._params_collection:
            p.IsSelected = False

    # ------------------------------------------------------------------
    # Target select all / none
    # ------------------------------------------------------------------

    def on_targets_select_all(self, sender, args):
        for t in self._targets_collection:
            t.IsSelected = True

    def on_targets_select_none(self, sender, args):
        for t in self._targets_collection:
            t.IsSelected = False

    # ------------------------------------------------------------------
    # OK / Cancel
    # ------------------------------------------------------------------

    def on_ok(self, sender, args):
        self.selected_params  = [p for p in self._all_params  if p.IsSelected]
        self.selected_targets = [t for t in self._target_items if t.IsSelected]
        self.selected_nested_families = [n for n in self._all_nested_items if n.IsSelected]
        self.selected_object_styles = [s for s in self._all_style_items if s.IsSelected]
        self.transfer_formulas = bool(self.chkCopyFormulas.IsChecked)
        self.transfer_values = bool(self.chkCopyValues.IsChecked)
        self.overwrite_existing_values = bool(self.chkOverwriteValues.IsChecked)
        self.create_missing_materials = bool(self.chkCreateMissingMaterials.IsChecked)
        self.nested_load_missing_only = bool(self.chkNestedLoadMissingOnly.IsChecked)
        self.nested_reload_existing = bool(self.chkNestedReloadExisting.IsChecked)

        if not self.selected_targets:
            forms.alert("No target families selected.", title="Transfer Parameters")
            return
        if not self.selected_params and not self.selected_nested_families and not self.selected_object_styles:
            forms.alert("No items selected to transfer in any tab.", title="Transfer Parameters")
            return

        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


# ---------------------------------------------------------------------------
# Collect source parameters
# ---------------------------------------------------------------------------

fm = doc.FamilyManager
all_param_items = []
for fp in fm.Parameters:
    try:
        item = ParamItem(fp)
        if _is_family_type_parameter(fp):
            item.FamilyTypeCategoryId = _detect_familytype_category_id(doc, fm, fp)
        all_param_items.append(item)
    except Exception as ex:
        output.print_md("**Warning:** Could not read parameter '{}': {}".format(
            fp.Definition.Name if fp.Definition else "?", ex))

all_param_items.sort(key=lambda p: (p.GroupLabel, p.Name))

# ---------------------------------------------------------------------------
# Collect open target family documents (exclude self)
# ---------------------------------------------------------------------------

target_items = []
for d in app.Documents:
    if d.IsFamilyDocument and d.Title != doc.Title:
        target_items.append(TargetItem(d))

if not target_items:
    forms.alert(
        "No other family documents are currently open.\n"
        "Open the target families in Revit first, then re-run this tool.",
        exitscript=True
    )

# ---------------------------------------------------------------------------
# Show dialog
# ---------------------------------------------------------------------------

nested_family_items = []
for fam_elem in sorted(
    FilteredElementCollector(doc).OfClass(Family).ToElements(),
    key=lambda f: f.Name
):
    try:
        n = (fam_elem.Name or "").lower()
        fam_cat = getattr(fam_elem, "FamilyCategory", None)
        cat_name = (fam_cat.Name or "").lower() if fam_cat is not None else ""

        # Exclude level heads and section marks/heads by both family name and category name.
        blocked_tokens = [
            "level head",
            "levelhead",
            "section mark",
            "sectionmark",
            "section head",
            "sectionhead",
        ]
        if any(tok in n for tok in blocked_tokens) or any(tok in cat_name for tok in blocked_tokens):
            continue
        nested_family_items.append(NestedFamilyItem(fam_elem))
    except Exception:
        pass

style_items = []
active_family_cat_name = None
try:
    active_family_cat = doc.OwnerFamily.FamilyCategory
    if active_family_cat is not None:
        active_family_cat_name = active_family_cat.Name
except Exception:
    active_family_cat_name = None

allowed_annotation_parent_names = {
    "reference lines",
    "reference line",
    "reference planes",
    "reference plane",
}

for parent_cat in doc.Settings.Categories:
    parent_name = (parent_cat.Name or "").strip()
    parent_name_l = parent_name.lower()

    # Include active family category plus specific annotation categories.
    include_active_family_cat = bool(active_family_cat_name and parent_name == active_family_cat_name)
    include_ref_annotation = (
        parent_name_l in allowed_annotation_parent_names
        and parent_cat.CategoryType == CategoryType.Annotation
    )

    if not include_active_family_cat and not include_ref_annotation:
        continue

    try:
        subcats = parent_cat.SubCategories
    except Exception:
        subcats = None
    if not subcats:
        continue

    for subcat in subcats:
        try:
            ctype = subcat.CategoryType
            if ctype == CategoryType.Model:
                stype = "Model"
            elif ctype == CategoryType.Annotation:
                stype = "Annotation"
            elif ctype == CategoryType.Import:
                stype = "Import"
            else:
                stype = str(ctype)

            style_items.append(ObjectStyleItem(parent_cat.Name, subcat.Name, stype, subcat))
        except Exception:
            pass

style_items.sort(key=lambda s: (s.StyleType, s.Name))

dialog = TransferParamsWindow(all_param_items, target_items, nested_family_items, style_items)
result = dialog.ShowDialog()

if not result:
    script.exit()

selected_params  = dialog.selected_params
selected_targets = dialog.selected_targets
selected_nested_families = dialog.selected_nested_families
selected_object_styles = dialog.selected_object_styles
transfer_formulas = dialog.transfer_formulas
transfer_values = dialog.transfer_values
overwrite_existing_values = dialog.overwrite_existing_values
create_missing_materials = dialog.create_missing_materials
nested_load_missing_only = dialog.nested_load_missing_only
nested_reload_existing = dialog.nested_reload_existing

# ---------------------------------------------------------------------------
# Execute transfers and report
# ---------------------------------------------------------------------------

output.print_md("# Parameter Transfer Results")
output.print_md("**Source family:** {}".format(doc.Title))
output.print_md("**Parameters selected:** {}".format(len(selected_params)))
output.print_md("**Nested families selected:** {}".format(len(selected_nested_families)))
output.print_md("**Object styles selected:** {}".format(len(selected_object_styles)))
output.print_md("**Transfer formulas:** {}".format("Yes" if transfer_formulas else "No"))
output.print_md("**Transfer values:** {}".format("Yes" if transfer_values else "No"))
output.print_md("**Overwrite existing values:** {}".format("Yes" if overwrite_existing_values else "No"))
output.print_md("**Create missing materials:** {}".format("Yes" if create_missing_materials else "No"))
output.print_md("---")

total_transferred = 0
total_skipped     = 0
total_failed      = 0

for target_item in selected_targets:
    t_doc = target_item.Document
    output.print_md("## Target: {}".format(t_doc.Title))

    if selected_nested_families:
        nested_res = transfer_nested_families_to_doc(
            doc,
            selected_nested_families,
            t_doc,
            nested_load_missing_only,
            nested_reload_existing,
        )

        if nested_res["loaded"]:
            output.print_md("**Nested families loaded ({}):**".format(len(nested_res["loaded"])))
            for name, _ in nested_res["loaded"]:
                output.print_md("- ✓ {}".format(name))
        if nested_res["reloaded"]:
            output.print_md("**Nested families reloaded ({}):**".format(len(nested_res["reloaded"])))
            for name, _ in nested_res["reloaded"]:
                output.print_md("- ✓ {}".format(name))
        if nested_res["skipped"]:
            output.print_md("**Nested families skipped ({}):**".format(len(nested_res["skipped"])))
            for name, reason in nested_res["skipped"]:
                output.print_md("- ⚬ {} &nbsp;—&nbsp; _{}_".format(name, reason))
        if nested_res["failed"]:
            output.print_md("**Nested families failed ({}):**".format(len(nested_res["failed"])))
            for name, reason in nested_res["failed"]:
                output.print_md("- ✗ {} &nbsp;—&nbsp; _{}_".format(name, reason))

    res = transfer_params_to_doc(
        selected_params,
        t_doc,
        transfer_formulas,
        transfer_values,
        overwrite_existing_values,
        create_missing_materials,
    )

    transferred = res["transferred"]
    skipped     = res["skipped"]
    failed      = res["failed"]

    if transferred:
        output.print_md("**Transferred ({}):**".format(len(transferred)))
        for name, _ in transferred:
            output.print_md("- ✓ {}".format(name))

    if skipped:
        output.print_md("**Skipped — already exists ({}):**".format(len(skipped)))
        for name, reason in skipped:
            output.print_md("- ⚬ {} &nbsp;—&nbsp; _{}_".format(name, reason))

    if failed:
        output.print_md("**Failed ({}):**".format(len(failed)))
        for name, reason in failed:
            output.print_md("- ✗ {} &nbsp;—&nbsp; _{}_".format(name, reason))

    formula_applied = res["formula_applied"]
    formula_failed  = res["formula_failed"]
    preload_notes   = res["preload_notes"]
    values_copied   = res["values_copied"]
    values_skipped  = res["values_skipped"]
    values_failed   = res["values_failed"]
    materials_created = res["materials_created"]
    if preload_notes:
        output.print_md("**Family Type preload notes ({}):**".format(len(preload_notes)))
        for note in preload_notes:
            output.print_md("- {}".format(note))
    if transfer_formulas:
        if formula_applied:
            output.print_md("**Formulas applied ({}):**".format(len(formula_applied)))
            for name, source_state in formula_applied:
                output.print_md("- ✓ {} &nbsp;—&nbsp; _from {} parameter_".format(name, source_state))
        if formula_failed:
            output.print_md("**Formula copy failed ({}):**".format(len(formula_failed)))
            for name, reason in formula_failed:
                output.print_md("- ✗ {} &nbsp;—&nbsp; _{}_".format(name, reason))

    if transfer_values:
        if materials_created:
            output.print_md("**Materials created ({}):**".format(len(materials_created)))
            for mat_name in materials_created:
                output.print_md("- ✓ {}".format(mat_name))
        if values_copied:
            output.print_md("**Values copied ({}):**".format(len(values_copied)))
            for name, source_state in values_copied:
                output.print_md("- ✓ {} &nbsp;—&nbsp; _from {} parameter_".format(name, source_state))
        if values_skipped:
            output.print_md("**Value copy skipped ({}):**".format(len(values_skipped)))
            for name, reason in values_skipped:
                output.print_md("- ⚬ {} &nbsp;—&nbsp; _{}_".format(name, reason))
        if values_failed:
            output.print_md("**Value copy failed ({}):**".format(len(values_failed)))
            for name, reason in values_failed:
                output.print_md("- ✗ {} &nbsp;—&nbsp; _{}_".format(name, reason))

    if selected_object_styles:
        style_res = transfer_object_styles_to_doc(selected_object_styles, t_doc)

        if style_res["created"]:
            output.print_md("**Object styles created ({}):**".format(len(style_res["created"])))
            for name, _ in style_res["created"]:
                output.print_md("- ✓ {}".format(name))
        if style_res["copied"]:
            output.print_md("**Object styles copied ({}):**".format(len(style_res["copied"])))
            for name, _ in style_res["copied"]:
                output.print_md("- ✓ {}".format(name))
        if style_res["skipped"]:
            output.print_md("**Object styles skipped ({}):**".format(len(style_res["skipped"])))
            for name, reason in style_res["skipped"]:
                output.print_md("- ⚬ {} &nbsp;—&nbsp; _{}_".format(name, reason))
        if style_res["failed"]:
            output.print_md("**Object styles failed ({}):**".format(len(style_res["failed"])))
            for name, reason in style_res["failed"]:
                output.print_md("- ✗ {} &nbsp;—&nbsp; _{}_".format(name, reason))

    total_transferred += len(transferred)
    total_skipped     += len(skipped)
    total_failed      += len(failed)
    output.print_md("---")

output.print_md(
    "**Summary across all targets:** "
    "{transferred} transferred, {skipped} skipped, {failed} failed.".format(
        transferred=total_transferred,
        skipped=total_skipped,
        failed=total_failed,
    )
)
