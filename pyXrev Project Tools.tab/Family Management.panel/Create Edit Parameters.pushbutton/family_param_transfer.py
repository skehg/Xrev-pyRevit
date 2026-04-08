# -*- coding: utf-8 -*-
"""
family_param_transfer.py
------------------------
Reusable library for transferring family parameters, nested families, and
object styles between open Revit family documents.

All functions take explicit ``app`` / ``source_doc`` arguments so they can be
called from any context (Parameter Editor tab, standalone script, etc.)
without relying on module-level globals.

Supports Revit 2020 – 2026 via ForgeTypeId / ParameterType dual-path.
"""

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Category,
    ElementId,
    Family,
    FamilySource,
    FamilySymbol,
    FilteredElementCollector,
    GraphicsStyleType,
    IFamilyLoadOptions,
    LabelUtils,
    Material,
    CategoryType,
    StorageType,
    SubTransaction,
    Transaction,
    ExternalDefinitionCreationOptions,
)

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs


# ---------------------------------------------------------------------------
# INotifyPropertyChanged helper mixin
# ---------------------------------------------------------------------------

class _NotifyBase(INotifyPropertyChanged):
    def __init__(self):
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


# ---------------------------------------------------------------------------
# Data model classes
# ---------------------------------------------------------------------------

class SourceParamItem(_NotifyBase):
    """One source family parameter for display in a WPF ListView."""

    def __init__(self, fp):
        _NotifyBase.__init__(self)
        self._fp = fp
        self.Name = fp.Definition.Name
        self.GroupLabel = _group_label(fp)
        self.InstanceTypeLabel = u"Instance" if fp.IsInstance else u"Type"
        self.Formula = fp.Formula or u""
        self.HasFormula = bool(self.Formula)
        self.IsShared = fp.IsShared
        self.SharedLabel = u"Yes" if fp.IsShared else u""
        self.IsLikelyNonTransferable = _is_likely_nontransferable(fp)
        # Non-transferable (built-in) params default to unselected
        self._is_selected = not self.IsLikelyNonTransferable
        try:
            self.Guid = fp.GUID
        except Exception:
            self.Guid = None
        self.FamilyTypeCategoryId = None

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify(u"IsSelected")

    @property
    def FamilyParameter(self):
        return self._fp


class SourceNestedFamilyItem(_NotifyBase):
    """One nested family element for display in a WPF ListView."""

    def __init__(self, family_elem):
        _NotifyBase.__init__(self)
        self._is_selected = False
        self._family = family_elem
        self.Name = family_elem.Name
        cat = getattr(family_elem, u"FamilyCategory", None)
        self.CategoryName = cat.Name if cat is not None else u"<No Category>"
        self.Status = u"Available"

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify(u"IsSelected")

    @property
    def Family(self):
        return self._family


class SourceObjectStyleItem(_NotifyBase):
    """One object style subcategory for display in a WPF ListView."""

    def __init__(self, parent_category_name, subcategory_name, style_type, source_subcategory):
        _NotifyBase.__init__(self)
        self._is_selected = False
        self.ParentCategory = parent_category_name
        self.Name = subcategory_name
        self.StyleType = style_type
        self.SourceSubcategory = source_subcategory
        self.Status = u"Available"

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify(u"IsSelected")


# ---------------------------------------------------------------------------
# Version-safe API helpers
# ---------------------------------------------------------------------------

def _get_data_type(definition):
    try:
        return definition.GetDataType()
    except AttributeError:
        return definition.ParameterType


def _group_of(family_param):
    """Return the parameter group as a ForgeTypeId (2023+) or BuiltInParameterGroup."""
    try:
        return family_param.Definition.GetGroupTypeId()
    except AttributeError:
        return family_param.Definition.ParameterGroup


def _group_label(family_param):
    g = _group_of(family_param)
    if g is None:
        return u""
    # Revit 2025+: ForgeTypeId groups require GetLabelForGroup
    try:
        label = LabelUtils.GetLabelForGroup(g)
        if label:
            return label
    except Exception:
        pass
    try:
        label = LabelUtils.GetLabelFor(g)
        if label:
            return label
    except Exception:
        pass
    return u""


def _is_likely_nontransferable(family_param):
    try:
        bip = family_param.Definition.BuiltInParameter
        if bip != BuiltInParameter.INVALID:
            return True
    except Exception:
        pass
    return False


def _is_family_type_parameter(family_param):
    try:
        ptype = family_param.Definition.ParameterType
        if str(ptype) == u"FamilyType":
            return True
    except Exception:
        pass
    try:
        data_type = family_param.Definition.GetDataType()
        if data_type:
            type_id = getattr(data_type, u"TypeId", u"")
            if type_id and u"autodesk.revit.category" in type_id.lower():
                return True
    except Exception:
        pass
    return False


def _try_get_familytype_category(target_doc, family_param):
    try:
        data_type = family_param.Definition.GetDataType()
    except Exception:
        data_type = None
    if data_type is None:
        return None
    try:
        cat = Category.GetCategory(target_doc, data_type)
        if cat:
            return cat
    except Exception:
        pass
    try:
        type_id = data_type.TypeId
        if type_id and u"ost_" in type_id.lower():
            token = type_id.split(u".")[-1]
            import System
            bic_type = System.Type.GetType(u"Autodesk.Revit.DB.BuiltInCategory, RevitAPI")
            if bic_type:
                bic = System.Enum.Parse(bic_type, token, True)
                return Category.GetCategory(target_doc, bic)
    except Exception:
        pass
    return None


def _category_by_integer_id(target_doc, category_int_id):
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
    if category is None:
        return False
    cat_id = category.Id.IntegerValue
    if cat_id in cache_dict:
        return cache_dict[cat_id]
    has_any = False
    try:
        types = (FilteredElementCollector(target_doc)
                 .WhereElementIsElementType()
                 .ToElements())
        for t in types:
            if t.Category is not None and t.Category.Id.IntegerValue == cat_id:
                has_any = True
                break
    except Exception:
        has_any = False
    cache_dict[cat_id] = has_any
    return has_any


def _collect_source_nested_families_for_param(source_doc, source_fm, family_param,
                                              expected_category_id=None):
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
            cat = getattr(elem, u"Category", None)
            if expected_category_id is not None:
                if cat is None or cat.Id.IntegerValue != int(expected_category_id):
                    continue
            fam = getattr(elem, u"Family", None)
            if fam is None:
                continue
            fam_map[fam.Id.IntegerValue] = fam
    except Exception:
        pass
    return list(fam_map.values())


class _FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = FamilySource.Family
        overwriteParameterValues.Value = True
        return True


def _preload_familytype_sources_into_target(source_doc, source_fm, target_doc,
                                            param_item, category_type_cache, results):
    fp = param_item.FamilyParameter
    name = param_item.Name
    cat = _category_by_integer_id(target_doc, param_item.FamilyTypeCategoryId)
    if cat is None:
        cat = _try_get_familytype_category(target_doc, fp)
    if cat is None:
        return
    if _target_has_types_for_category(target_doc, cat, category_type_cache):
        return
    source_nested = _collect_source_nested_families_for_param(
        source_doc, source_fm, fp,
        expected_category_id=param_item.FamilyTypeCategoryId)
    if not source_nested:
        results[u"preload_notes"].append(
            u"{}: no source nested families found for category '{}'".format(name, cat.Name))
        return
    load_opts = _FamilyLoadOptions()
    loaded = 0
    for nf in source_nested:
        nd = None
        try:
            nd = source_doc.EditFamily(nf)
            nd.LoadFamily(target_doc, load_opts)
            loaded += 1
        except Exception as ex:
            results[u"preload_notes"].append(
                u"{}: could not preload '{}': {}".format(name, nf.Name, ex))
        finally:
            if nd:
                try:
                    nd.Close(False)
                except Exception:
                    pass
    category_type_cache.pop(cat.Id.IntegerValue, None)
    if loaded:
        results[u"preload_notes"].append(
            u"{}: preloaded {} family/families for '{}'".format(name, loaded, cat.Name))


# ---------------------------------------------------------------------------
# Shared parameter file helpers
# ---------------------------------------------------------------------------

def _open_def_file(app):
    try:
        return app.OpenSharedParameterFile()
    except Exception:
        return None


def _find_definition_by_guid(def_file, guid):
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


def ensure_in_sharedparam_file(app, param_item):
    """Locate or create the shared parameter definition in the active .txt file.

    Returns ExternalDefinition or None if no shared parameter file is configured.
    """
    def_file = _open_def_file(app)
    if def_file is None:
        return None
    existing = _find_definition_by_guid(def_file, param_item.Guid)
    if existing:
        return existing
    fp = param_item.FamilyParameter
    data_type = _get_data_type(fp.Definition)
    export_group = def_file.Groups[u"pyEXPORT"]
    if export_group is None:
        export_group = def_file.Groups.Create(u"pyEXPORT")
    opts = ExternalDefinitionCreationOptions(param_item.Name, data_type)
    opts.GUID = param_item.Guid
    opts.Visible = True
    opts.UserModifiable = True
    try:
        opts.Description = fp.Definition.Description
    except Exception:
        pass
    return export_group.Definitions.Create(opts)


# ---------------------------------------------------------------------------
# Parameter transfer helpers
# ---------------------------------------------------------------------------

def _existing_param_names(family_manager):
    names = set()
    for fp in family_manager.Parameters:
        names.add(fp.Definition.Name)
    return names


def _find_param_by_name(family_manager, name):
    for fp in family_manager.Parameters:
        if fp.Definition.Name == name:
            return fp
    return None


def _safe_elem_name(elem):
    if elem is None:
        return None
    try:
        if elem.Name:
            return elem.Name
    except Exception:
        pass
    for bip in [BuiltInParameter.SYMBOL_NAME_PARAM, BuiltInParameter.ALL_MODEL_TYPE_NAME]:
        try:
            p = elem.get_Parameter(bip)
            if p:
                n = p.AsString()
                if n:
                    return n
        except Exception:
            pass
    return None


def _safe_elem_family_name(elem):
    if elem is None:
        return None
    try:
        if elem.FamilyName:
            return elem.FamilyName
    except Exception:
        pass
    try:
        fam = getattr(elem, u"Family", None)
        if fam and fam.Name:
            return fam.Name
    except Exception:
        pass
    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if p:
            n = p.AsString()
            if n:
                return n
    except Exception:
        pass
    return None


def _safe_category_name(cat):
    if cat is None:
        return u"<No Category>"
    try:
        if cat.Name:
            return cat.Name
    except Exception:
        pass
    try:
        return u"CategoryId {}".format(cat.Id.IntegerValue)
    except Exception:
        return u"<Unknown Category>"


def _get_or_build_type_cache(target_doc, cache_dict):
    key, key_f = u"_types_by_cat_and_name", u"_types_by_cat_family_and_name"
    if key in cache_dict and key_f in cache_dict:
        return cache_dict[key], cache_dict[key_f]
    mapping, mapping_f = {}, {}
    try:
        for et in (FilteredElementCollector(target_doc)
                   .WhereElementIsElementType()
                   .ToElements()):
            cat = getattr(et, u"Category", None)
            if cat is None:
                continue
            tn = _safe_elem_name(et)
            if not tn:
                continue
            k = (cat.Id.IntegerValue, tn)
            if k not in mapping:
                mapping[k] = et
            fn = _safe_elem_family_name(et)
            if fn:
                kf = (cat.Id.IntegerValue, fn, tn)
                if kf not in mapping_f:
                    mapping_f[kf] = et
    except Exception:
        pass
    cache_dict[key] = mapping
    cache_dict[key_f] = mapping_f
    return mapping, mapping_f


def _get_or_build_symbol_cache(target_doc, cache_dict):
    key, key_f = u"_symbols_by_cat_and_name", u"_symbols_by_cat_family_and_name"
    if key in cache_dict and key_f in cache_dict:
        return cache_dict[key], cache_dict[key_f]
    mapping, mapping_f = {}, {}
    try:
        for sym in FilteredElementCollector(target_doc).OfClass(FamilySymbol).ToElements():
            cat = getattr(sym, u"Category", None)
            if cat is None:
                continue
            tn = _safe_elem_name(sym)
            if not tn:
                continue
            k = (cat.Id.IntegerValue, tn)
            if k not in mapping:
                mapping[k] = sym
            fn = _safe_elem_family_name(sym)
            if fn:
                kf = (cat.Id.IntegerValue, fn, tn)
                if kf not in mapping_f:
                    mapping_f[kf] = sym
    except Exception:
        pass
    cache_dict[key] = mapping
    cache_dict[key_f] = mapping_f
    return mapping, mapping_f


def _get_or_build_material_cache(target_doc, cache_dict):
    key = u"_materials_by_name"
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
    key = u"_created_material_names"
    if key not in cache_dict:
        cache_dict[key] = set()
    return cache_dict[key]


def _copy_material_properties(source_mat, target_mat):
    if source_mat is None or target_mat is None:
        return
    for prop in [u"Color", u"Transparency", u"Shininess", u"Smoothness",
                 u"UseRenderAppearanceForShading",
                 u"CutForegroundPatternColor", u"CutForegroundPatternId",
                 u"CutBackgroundPatternColor", u"CutBackgroundPatternId",
                 u"SurfaceForegroundPatternColor", u"SurfaceForegroundPatternId",
                 u"SurfaceBackgroundPatternColor", u"SurfaceBackgroundPatternId",
                 u"AppearanceAssetId"]:
        try:
            if hasattr(source_mat, prop) and hasattr(target_mat, prop):
                setattr(target_mat, prop, getattr(source_mat, prop))
        except Exception:
            pass


def _ensure_material_in_target(target_doc, mat_name, cache_dict, allow_create,
                                source_material=None):
    mats = _get_or_build_material_cache(target_doc, cache_dict)
    target_mat = mats.get(mat_name)
    if target_mat:
        return True, target_mat, u""
    if not allow_create:
        return False, None, u"Material '{}' not found in target".format(mat_name)
    try:
        new_id = Material.Create(target_doc, mat_name)
        new_mat = target_doc.GetElement(new_id)
        if new_mat is None:
            return False, None, u"Material '{}' could not be created".format(mat_name)
        _copy_material_properties(source_material, new_mat)
        mats[mat_name] = new_mat
        _get_or_build_created_materials(cache_dict).add(mat_name)
        return True, new_mat, u""
    except Exception as ex:
        return False, None, u"Material '{}' creation failed: {}".format(mat_name, ex)


def _map_element_id_value(source_doc, target_doc, source_elem_id, cache_dict,
                          preferred_type_name=None, require_family_symbol=False,
                          create_missing_materials=False):
    if source_elem_id is None or source_elem_id == ElementId.InvalidElementId:
        return True, ElementId.InvalidElementId, u""
    source_elem = source_doc.GetElement(source_elem_id)
    if source_elem is None:
        return False, None, u"Source ElementId does not resolve to an element"
    if isinstance(source_elem, Material):
        ok, mat, reason = _ensure_material_in_target(
            target_doc, source_elem.Name, cache_dict, create_missing_materials, source_elem)
        if ok:
            return True, mat.Id, u""
        return False, None, reason
    cat = getattr(source_elem, u"Category", None)
    if cat is not None:
        stn = _safe_elem_name(source_elem) or preferred_type_name
        if not stn:
            return False, None, u"Could not read source type name for ElementId mapping"
        sfn = _safe_elem_family_name(source_elem)
        if require_family_symbol:
            tm, tfm = _get_or_build_symbol_cache(target_doc, cache_dict)
        else:
            tm, tfm = _get_or_build_type_cache(target_doc, cache_dict)
        target_type = None
        if sfn:
            target_type = tfm.get((cat.Id.IntegerValue, sfn, stn))
        if target_type is None and sfn and preferred_type_name and preferred_type_name != stn:
            target_type = tfm.get((cat.Id.IntegerValue, sfn, preferred_type_name))
        if target_type is None:
            target_type = tm.get((cat.Id.IntegerValue, stn))
        if target_type is None and preferred_type_name and preferred_type_name != stn:
            target_type = tm.get((cat.Id.IntegerValue, preferred_type_name))
        if target_type:
            return True, target_type.Id, u""
        return False, None, (
            u"Could not map type '{}' (family '{}') in category '{}'".format(
                stn, sfn or u"?", _safe_category_name(cat)))
    return False, None, u"Unsupported ElementId value mapping"


def _copy_value_from_source_to_target(source_doc, source_fm, source_fp,
                                      target_doc, target_fm, target_fp, cache_dict,
                                      create_missing_materials=False):
    if source_fp.Formula:
        return False, u"Skipped: source parameter has formula"
    source_type = source_fm.CurrentType
    target_type = target_fm.CurrentType
    if source_type is None:
        return False, u"Source family has no current type"
    if target_type is None:
        return False, u"Target family has no current type"
    st = source_fp.StorageType
    try:
        if st == StorageType.Double:
            target_fm.Set(target_fp, source_type.AsDouble(source_fp))
        elif st == StorageType.Integer:
            target_fm.Set(target_fp, source_type.AsInteger(source_fp))
        elif st == StorageType.String:
            val = source_type.AsString(source_fp)
            target_fm.Set(target_fp, val if val is not None else u"")
        elif st == StorageType.ElementId:
            preferred = None
            try:
                preferred = source_type.AsValueString(source_fp)
            except Exception:
                pass
            req_sym = False
            try:
                req_sym = _is_family_type_parameter(target_fp)
            except Exception:
                pass
            ok, mapped_id, reason = _map_element_id_value(
                source_doc, target_doc, source_type.AsElementId(source_fp),
                cache_dict, preferred, req_sym, create_missing_materials)
            if not ok:
                return False, reason
            if req_sym:
                me = target_doc.GetElement(mapped_id)
                if not isinstance(me, FamilySymbol):
                    return False, u"Mapped value is not a FamilySymbol"
            target_fm.Set(target_fp, mapped_id)
        else:
            return False, u"Unsupported storage type: {}".format(st)
    except Exception as ex:
        return False, str(ex)
    return True, u""


# ---------------------------------------------------------------------------
# Public transfer functions
# ---------------------------------------------------------------------------

def transfer_params_to_doc(app, source_doc, param_items, target_doc,
                           transfer_formulas=True, transfer_values=True,
                           overwrite_existing_values=False,
                           create_missing_materials=False):
    """Transfer SourceParamItems from source_doc into target_doc's FamilyManager.

    Returns a results dict with keys:
        'transferred', 'skipped', 'failed',
        'formula_applied', 'formula_failed',
        'values_copied', 'values_skipped', 'values_failed',
        'materials_created', 'preload_notes'
    Each value is a list of (name, reason) tuples (reason empty string = success).
    """
    results = {
        u"transferred": [], u"skipped": [], u"failed": [],
        u"formula_applied": [], u"formula_failed": [],
        u"values_copied": [], u"values_skipped": [], u"values_failed": [],
        u"materials_created": [], u"preload_notes": [],
    }
    fm = target_doc.FamilyManager
    source_fm = source_doc.FamilyManager
    existing = _existing_param_names(fm)
    formula_candidates = []
    value_candidates = []
    category_type_cache = {}
    value_map_cache = {}

    t = Transaction(target_doc, u"Import Family Parameters")
    try:
        for item in param_items:
            if item.Name in existing:
                continue
            if _is_family_type_parameter(item.FamilyParameter):
                _preload_familytype_sources_into_target(
                    source_doc, source_fm, target_doc, item, category_type_cache, results)

        t.Start()
        for item in param_items:
            name = item.Name
            if name in existing:
                results[u"skipped"].append((name, u"Already exists in target"))
                if transfer_formulas and item.Formula:
                    formula_candidates.append((item, u"existing"))
                if transfer_values:
                    if overwrite_existing_values:
                        value_candidates.append((item, u"existing"))
                    else:
                        results[u"values_skipped"].append(
                            (name, u"Existing parameter and overwrite is disabled"))
                continue
            try:
                st = SubTransaction(target_doc)
                st.Start()
                if item.IsShared:
                    ext_defn = ensure_in_sharedparam_file(app, item)
                    if ext_defn is None:
                        results[u"failed"].append(
                            (name, u"No shared parameter file configured"))
                        st.RollBack()
                        continue
                    fm.AddParameter(ext_defn, _group_of(item.FamilyParameter),
                                    item.FamilyParameter.IsInstance)
                else:
                    fp = item.FamilyParameter
                    grp = _group_of(fp)
                    is_inst = fp.IsInstance
                    data_type = _get_data_type(fp.Definition)
                    try:
                        fm.AddParameter(name, grp, data_type, is_inst)
                    except Exception as add_ex:
                        if _is_family_type_parameter(fp):
                            cat = _category_by_integer_id(target_doc, item.FamilyTypeCategoryId)
                            if cat is None:
                                cat = _try_get_familytype_category(target_doc, fp)
                            if cat is None:
                                raise Exception(
                                    u"Family Type parameter requires category overload; "
                                    u"could not resolve category")
                            if not _target_has_types_for_category(target_doc, cat, category_type_cache):
                                raise Exception(
                                    u"Target has no types in category '{}' for this Family Type "
                                    u"parameter".format(cat.Name))
                            fm.AddParameter(name, grp, cat, is_inst)
                        else:
                            raise add_ex
                target_doc.Regenerate()
                st.Commit()
                results[u"transferred"].append((name, u""))
                existing.add(name)
                if transfer_formulas and item.Formula:
                    formula_candidates.append((item, u"transferred"))
                if transfer_values:
                    value_candidates.append((item, u"transferred"))
            except Exception as ex:
                try:
                    st.RollBack()
                except Exception:
                    pass
                results[u"failed"].append((name, str(ex)))

        t.Commit()

        if transfer_values and value_candidates:
            tv = Transaction(target_doc, u"Import Parameter Values")
            try:
                tv.Start()
                for item, src_state in value_candidates:
                    target_fp = _find_param_by_name(fm, item.Name)
                    if target_fp is None:
                        results[u"values_failed"].append(
                            (item.Name, u"Could not find matching parameter in target"))
                        continue
                    ok, reason = _copy_value_from_source_to_target(
                        source_doc, source_fm, item.FamilyParameter,
                        target_doc, fm, target_fp, value_map_cache, create_missing_materials)
                    if ok:
                        results[u"values_copied"].append((item.Name, src_state))
                    else:
                        if reason.startswith(u"Skipped:"):
                            results[u"values_skipped"].append(
                                (item.Name, reason.replace(u"Skipped: ", u"")))
                        else:
                            results[u"values_failed"].append((item.Name, reason))
                tv.Commit()
                created = sorted(list(_get_or_build_created_materials(value_map_cache)))
                if created:
                    results[u"materials_created"] = created
            except Exception as vex:
                try:
                    tv.RollBack()
                except Exception:
                    pass
                rec = ({n for n, _ in results[u"values_copied"]} |
                       {n for n, _ in results[u"values_failed"]} |
                       {n for n, _ in results[u"values_skipped"]})
                for item, _ in value_candidates:
                    if item.Name not in rec:
                        results[u"values_failed"].append(
                            (item.Name, u"Value transaction rolled back: " + str(vex)))

        if transfer_formulas and formula_candidates:
            tf = Transaction(target_doc, u"Import Parameter Formulas")
            try:
                tf.Start()
                for item, src_state in formula_candidates:
                    target_fp = _find_param_by_name(fm, item.Name)
                    if target_fp is None:
                        results[u"formula_failed"].append(
                            (item.Name, u"Could not find matching parameter in target"))
                        continue
                    try:
                        fm.SetFormula(target_fp, item.Formula)
                        results[u"formula_applied"].append((item.Name, src_state))
                    except Exception as fex:
                        results[u"formula_failed"].append((item.Name, str(fex)))
                tf.Commit()
            except Exception as ftex:
                try:
                    tf.RollBack()
                except Exception:
                    pass
                rec_a = {n for n, _ in results[u"formula_applied"]}
                rec_f = {n for n, _ in results[u"formula_failed"]}
                for item, _ in formula_candidates:
                    if item.Name not in rec_a and item.Name not in rec_f:
                        results[u"formula_failed"].append(
                            (item.Name, u"Formula transaction rolled back: " + str(ftex)))

    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        done = ({n for n, _ in results[u"transferred"]} |
                {n for n, _ in results[u"skipped"]} |
                {n for n, _ in results[u"failed"]})
        for item in param_items:
            if item.Name not in done:
                results[u"failed"].append((item.Name, u"Transaction rolled back: " + str(ex)))

    return results


def transfer_nested_families_to_doc(source_doc, nested_items, target_doc,
                                    load_missing_only=True, reload_existing=False):
    """Transfer nested families from source_doc into target_doc.

    Returns a results dict with keys: 'loaded', 'reloaded', 'skipped', 'failed'.
    """
    results = {u"loaded": [], u"reloaded": [], u"skipped": [], u"failed": []}
    target_family_names = set()
    try:
        for fam in FilteredElementCollector(target_doc).OfClass(Family).ToElements():
            target_family_names.add(fam.Name)
    except Exception:
        pass
    load_opts = _FamilyLoadOptions()
    for item in nested_items:
        fam = item.Family
        name = item.Name
        exists = name in target_family_names
        if exists and load_missing_only and not reload_existing:
            results[u"skipped"].append((name, u"Already exists in target"))
            continue
        nd = None
        try:
            nd = source_doc.EditFamily(fam)
            nd.LoadFamily(target_doc, load_opts)
            if exists:
                results[u"reloaded"].append((name, u""))
            else:
                results[u"loaded"].append((name, u""))
                target_family_names.add(name)
        except Exception as ex:
            results[u"failed"].append((name, str(ex)))
        finally:
            if nd:
                try:
                    nd.Close(False)
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
        for sc in parent_cat.SubCategories:
            if sc.Name == subcat_name:
                return sc
    except Exception:
        pass
    return None


def _copy_subcategory_graphics(source_sc, target_sc):
    try:
        target_sc.LineColor = source_sc.LineColor
    except Exception:
        pass
    try:
        target_sc.Material = source_sc.Material
    except Exception:
        pass
    for gst in [GraphicsStyleType.Projection, GraphicsStyleType.Cut]:
        try:
            lw = source_sc.GetLineWeight(gst)
            if lw and lw > 0:
                target_sc.SetLineWeight(lw, gst)
        except Exception:
            pass
        try:
            lp = source_sc.GetLinePatternId(gst)
            if lp:
                target_sc.SetLinePatternId(lp, gst)
        except Exception:
            pass


def transfer_object_styles_to_doc(source_doc, style_items, target_doc):
    """Transfer object style subcategories from source_doc to target_doc.

    Returns a results dict with keys: 'copied', 'created', 'skipped', 'failed'.
    """
    results = {u"copied": [], u"created": [], u"skipped": [], u"failed": []}
    tx = Transaction(target_doc, u"Import Object Styles")
    try:
        tx.Start()
        for item in style_items:
            label = u"{} : {}".format(item.ParentCategory, item.Name)
            src_sc = item.SourceSubcategory
            if src_sc is None:
                results[u"failed"].append((label, u"Source subcategory unavailable"))
                continue
            parent_cat = _find_parent_category_by_name(target_doc, item.ParentCategory)
            if parent_cat is None:
                results[u"skipped"].append((label, u"Parent category not found in target"))
                continue
            target_sc = _find_subcategory_by_name(parent_cat, item.Name)
            if target_sc is None:
                try:
                    is_sys = item.SourceSubcategory.Id.IntegerValue < 0
                except Exception:
                    is_sys = False
                if is_sys:
                    results[u"skipped"].append(
                        (label, u"System subcategory missing in target; cannot create"))
                    continue
                try:
                    target_sc = target_doc.Settings.Categories.NewSubcategory(
                        parent_cat, item.Name)
                    results[u"created"].append((label, u""))
                except Exception as ex:
                    results[u"failed"].append(
                        (label, u"Could not create subcategory: {}".format(ex)))
                    continue
            try:
                _copy_subcategory_graphics(src_sc, target_sc)
                results[u"copied"].append((label, u""))
            except Exception as ex:
                results[u"failed"].append((label, str(ex)))
        tx.Commit()
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        done = ({n for n, _ in results[u"copied"]} | {n for n, _ in results[u"created"]} |
                {n for n, _ in results[u"skipped"]} | {n for n, _ in results[u"failed"]})
        for item in style_items:
            label = u"{} : {}".format(item.ParentCategory, item.Name)
            if label not in done:
                results[u"failed"].append((label, u"Transaction rolled back: " + str(ex)))
    return results


# ---------------------------------------------------------------------------
# Discovery utilities
# ---------------------------------------------------------------------------

def get_open_family_docs(app, exclude_title=None):
    """Return list of open family Documents, optionally excluding one by title."""
    docs = []
    for d in app.Documents:
        if d.IsFamilyDocument:
            if exclude_title is None or d.Title != exclude_title:
                docs.append(d)
    return sorted(docs, key=lambda d: d.Title)


def read_params_from_doc(source_doc):
    """Read all family parameters from source_doc.

    Returns ObservableCollection[SourceParamItem] sorted by group then name.
    """
    col = ObservableCollection[object]()
    fm = source_doc.FamilyManager
    items = []
    for fp in fm.Parameters:
        try:
            item = SourceParamItem(fp)
            if _is_family_type_parameter(fp):
                item.FamilyTypeCategoryId = _detect_familytype_category_id(source_doc, fm, fp)
            items.append(item)
        except Exception:
            pass
    items.sort(key=lambda i: (i.GroupLabel.lower(), i.Name.lower()))
    for item in items:
        col.Add(item)
    return col


def read_nested_families_from_doc(source_doc):
    """Return ObservableCollection[SourceNestedFamilyItem] from source_doc, sorted by name.

    Filters out level heads and section marks (same heuristic as Transfer Parameters tool).
    """
    col = ObservableCollection[object]()
    blocked = [u"level head", u"levelhead", u"section mark",
               u"sectionmark", u"section head", u"sectionhead"]
    items = []
    for fam in sorted(
        FilteredElementCollector(source_doc).OfClass(Family).ToElements(),
        key=lambda f: f.Name
    ):
        try:
            n = (fam.Name or u"").lower()
            fam_cat = getattr(fam, u"FamilyCategory", None)
            cn = (fam_cat.Name or u"").lower() if fam_cat else u""
            if any(tok in n for tok in blocked) or any(tok in cn for tok in blocked):
                continue
            items.append(SourceNestedFamilyItem(fam))
        except Exception:
            pass
    for item in items:
        col.Add(item)
    return col


def read_object_styles_from_doc(source_doc, include_model=True,
                                include_annotation=True, include_imports=False,
                                parent_category_name=None):
    """Return ObservableCollection[SourceObjectStyleItem] from source_doc.

    Filters by style type flags and optional parent category name.
    """
    col = ObservableCollection[object]()
    active_family_cat_name = None
    try:
        afc = source_doc.OwnerFamily.FamilyCategory
        if afc is not None:
            active_family_cat_name = afc.Name
    except Exception:
        pass

    allowed_anno = {u"reference lines", u"reference line",
                    u"reference planes", u"reference plane"}

    items = []
    for parent_cat in source_doc.Settings.Categories:
        pname = (parent_cat.Name or u"").strip()
        pname_l = pname.lower()
        inc_family = bool(active_family_cat_name and pname == active_family_cat_name)
        inc_ref = (pname_l in allowed_anno and
                   parent_cat.CategoryType == CategoryType.Annotation)
        if not inc_family and not inc_ref:
            continue
        if parent_category_name is not None and pname != parent_category_name:
            continue
        try:
            subcats = parent_cat.SubCategories
        except Exception:
            subcats = None
        if not subcats:
            continue
        for sc in subcats:
            try:
                ct = sc.CategoryType
                if ct == CategoryType.Model:
                    stype = u"Model"
                    if not include_model:
                        continue
                elif ct == CategoryType.Annotation:
                    stype = u"Annotation"
                    if not include_annotation:
                        continue
                elif ct == CategoryType.Import:
                    stype = u"Import"
                    if not include_imports:
                        continue
                else:
                    stype = str(ct)
                items.append(SourceObjectStyleItem(pname, sc.Name, stype, sc))
            except Exception:
                pass

    items.sort(key=lambda s: (s.StyleType, s.Name))
    for item in items:
        col.Add(item)
    return col
