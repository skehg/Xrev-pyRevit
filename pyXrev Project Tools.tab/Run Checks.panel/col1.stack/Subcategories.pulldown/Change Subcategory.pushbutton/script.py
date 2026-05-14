# -*- coding: utf-8 -*-
# Reassign Family Geometry from Old Subcategory to New Subcategory
# Reads a CSV with: Family Name, Subcategory Name, New Subcategory
# Cleans families, saves to user-selected folder, reloads into project,
# purges matching project-level subcategories, and writes a timestamped audit CSV with parent category info.
#
# Updated: Fix reassignment when a family contains multiple source subcategories
# that should be moved (including multiple sources mapping to the same target).
# The reassignment now builds a direct lookup from source subcategory Id -> (old_name, target_cat)
# so elements are correctly detected and moved even when multiple subcategories exist.

import csv
import os
import codecs
import traceback
import datetime
import re

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    Transaction,
    BuiltInParameter,
    IFamilyLoadOptions,
    ElementId,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, revit
from pyrevit.forms import pick_file, pick_folder

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# --- Suppress overwrite dialogs when reloading families ---
class SimpleFamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        # instruct Revit to overwrite existing family when reloading
        overwriteParameterValues = True
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues = True
        return True

def safe_filename(name):
    # Replace characters not allowed in filenames
    return re.sub(r'[<>:"/\\|?*\n\r]+', '_', name).strip() or "family"

# --- Select CSV file ---
csv_path = pick_file(file_ext='csv', title='Select Subcategory Mapping CSV')
if not csv_path:
    TaskDialog.Show("Cancelled", "No CSV selected. Script aborted.")
    script.exit()

# --- Select output folder for processed families ---
save_folder = pick_folder(title="Select folder to save processed families and audit log")
if not save_folder:
    TaskDialog.Show("Cancelled", "No folder selected. Script aborted.")
    script.exit()

# --- Read CSV mappings and group by family ---
mappings_by_family = {}
mappings_list = []
try:
    with codecs.open(csv_path, "r", encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        expected = {"Family Name", "Subcategory Name", "New Subcategory"}
        if not expected.issubset(set(reader.fieldnames or [])):
            TaskDialog.Show("Error", "CSV headers must include: Family Name, Subcategory Name, New Subcategory")
            script.exit()

        for row in reader:
            fam_name = (row.get("Family Name") or "").strip()
            old_subcat = (row.get("Subcategory Name") or "").strip()
            new_subcat = (row.get("New Subcategory") or "").strip()
            if not (fam_name and old_subcat and new_subcat):
                continue
            mappings_list.append((fam_name, old_subcat, new_subcat))
            mappings_by_family.setdefault(fam_name, []).append((old_subcat, new_subcat))
except Exception:
    output.print_md("❌ Error reading CSV:\n{}".format(traceback.format_exc()))
    script.exit()

if not mappings_list:
    TaskDialog.Show("Error", "No valid mappings found in CSV.")
    script.exit()

# --- Collect families in project once ---
fam_collector = FilteredElementCollector(doc).OfClass(Family)
fam_dict = {fam.Name: fam for fam in fam_collector}

# --- Tracking summary ---
total_families = 0
total_moved = 0

deleted_subcats_family = []   # (family, old_subcat, parent)
failed_deletes_family = []    # (family, old_subcat, parent, reason)

deleted_subcats_project = []  # (project_subcat_name, parent)
failed_deletes_project = []   # (project_subcat_name, parent, reason)

family_results = []  # detailed per-family records for CSV

# --- Process each family once (grouped mappings) ---
for fam_name, mapping_pairs in mappings_by_family.items():
    if fam_name not in fam_dict:
        output.print_md("⚠️ Family '{}' not found in project.".format(fam_name))
        # add entries for each mapping row for auditing
        for old_subcat, new_subcat in mapping_pairs:
            family_results.append({
                "family": fam_name, "old": old_subcat, "new": new_subcat,
                "parent": "", "moved": 0, "old_deleted": False,
                "delete_reason": "Family not in project", "saved_path": ""
            })
        continue

    fam = fam_dict[fam_name]
    try:
        fam_doc = doc.EditFamily(fam)
        if not fam_doc:
            output.print_md("❌ Could not open family '{}' for editing.".format(fam_name))
            continue

        root_cat = fam_doc.OwnerFamily.FamilyCategory
        parent_name = root_cat.Name

        # Build mapping dictionaries for quick lookup: old_name -> new_name
        combined_map = {}
        for old_subcat, new_subcat in mapping_pairs:
            if old_subcat == new_subcat:
                # nothing to do for identical mapping
                continue
            combined_map.setdefault(old_subcat, new_subcat)

        if not combined_map:
            output.print_md("ℹ️ No meaningful subcategory mappings for family '{}'.".format(fam_name))
            fam_doc.Close(False)
            # write audit entries with no action
            for old_subcat, new_subcat in mapping_pairs:
                family_results.append({
                    "family": fam_name, "old": old_subcat, "new": new_subcat,
                    "parent": parent_name, "moved": 0, "old_deleted": False,
                    "delete_reason": "No action required (old == new)", "saved_path": ""
                })
            continue

        # Prepare targets: ensure all target subcategories exist (create if needed)
        target_subcats = {}  # new_name -> Category object
        for new_name in set(combined_map.values()):
            try:
                target = root_cat.SubCategories.get_Item(new_name)
            except:
                target = None
            if not target:
                # create inside transaction below
                target_subcats[new_name] = None
            else:
                target_subcats[new_name] = target

        # Resolve source subcategory objects (if present)
        source_subcats = {}  # old_name -> Category object
        for old_name in combined_map.keys():
            try:
                source = root_cat.SubCategories.get_Item(old_name)
            except:
                source = None
            source_subcats[old_name] = source

        # One transaction to create missing targets, reassign element subcategories, and attempt deletes
        tx = Transaction(fam_doc, "Reassign subcategories and cleanup")
        tx.Start()

        moved_counts = {old: 0 for old in combined_map.keys()}

        # Create missing target subcategories first
        for new_name, cat in list(target_subcats.items()):
            if cat is None:
                try:
                    target = fam_doc.Settings.Categories.NewSubcategory(root_cat, new_name)
                    target_subcats[new_name] = target
                    output.print_md("ℹ️ Created new subcategory '{}' in family '{}'.".format(new_name, fam_name))
                except Exception as e:
                    output.print_md("⚠️ Failed to create subcategory '{}' in family '{}': {}".format(new_name, fam_name, str(e)))
                    target_subcats[new_name] = None  # keep as None so reassign won't target it

        # Build a fast lookup: source_subcat_id.IntegerValue -> (old_name, target_category)
        id_lookup = {}
        for old_name, source_cat in source_subcats.items():
            if source_cat is None:
                continue
            new_name = combined_map.get(old_name)
            target_cat = target_subcats.get(new_name)
            if target_cat is None:
                # if there is no valid target, skip mapping for this source
                continue
            try:
                id_lookup[source_cat.Id.IntegerValue] = (old_name, target_cat)
            except Exception:
                # defensive: skip if Id not available for some reason
                continue

        # Reassign geometry elements that belong to any of the source subcategories using the id_lookup
        collector = FilteredElementCollector(fam_doc).WhereElementIsNotElementType()
        for e in collector:
            try:
                p = e.get_Parameter(BuiltInParameter.FAMILY_ELEM_SUBCATEGORY)
                if not p:
                    continue
                elem_subcat_id = p.AsElementId()
                if not elem_subcat_id:
                    continue
                key = elem_subcat_id.IntegerValue
                if key in id_lookup:
                    old_name, target_cat = id_lookup[key]
                    # perform the set
                    p.Set(target_cat.Id)
                    moved_counts[old_name] += 1
            except:
                # element may not allow parameter set; ignore and continue
                continue

        # Attempt to delete each source subcategory (only after reassign attempts)
        old_deleted_map = {}
        delete_reasons = {}
        for old_name, source_cat in source_subcats.items():
            if source_cat is None:
                old_deleted_map[old_name] = False
                delete_reasons[old_name] = "Source subcategory not present in family"
                continue
            try:
                fam_doc.Delete(source_cat.Id)
                old_deleted_map[old_name] = True
                deleted_subcats_family.append((fam_name, old_name, parent_name))
                output.print_md("🗑️ Deleted old subcategory '{}' (parent: '{}') from family '{}'.".format(
                    old_name, parent_name, fam_name))
            except Exception as del_err:
                old_deleted_map[old_name] = False
                delete_reasons[old_name] = str(del_err)
                failed_deletes_family.append((fam_name, old_name, parent_name, str(del_err)))
                output.print_md("⚠️ Could not delete subcategory '{}' (parent: '{}') in family '{}': {}".format(
                    old_name, parent_name, fam_name, str(del_err)))

        tx.Commit()

        # Save once to chosen folder (sanitized filename) and reload once
        safe_name = safe_filename(fam_name)
        save_path = os.path.join(save_folder, "{}.rfa".format(safe_name))
        # remove existing file to avoid prompts
        try:
            if os.path.exists(save_path):
                try:
                    os.remove(save_path)
                except Exception:
                    # if we can't delete, we'll let SaveAs attempt and handle potential prompts
                    pass
            fam_doc.SaveAs(save_path)
        except Exception:
            # Fallback: Save to original path (if available)
            try:
                fam_doc.Save()
                save_path = fam_doc.PathName or save_path
            except Exception:
                # give up save but still attempt to close and continue
                output.print_md("⚠️ Failed to SaveAs or Save family '{}'. It may not be reloaded.".format(fam_name))

        # Close family doc without an extra save (we already saved via SaveAs above)
        try:
            fam_doc.Close(False)
        except:
            # best-effort close
            try:
                fam_doc.Close(True)
            except:
                pass

        # Reload edited family into project (single reload)
        try:
            tx_reload = Transaction(doc, "Reload Family '{}'".format(fam_name))
            tx_reload.Start()
            # Use load options to overwrite existing family
            doc.LoadFamily(save_path, SimpleFamilyLoadOptions())
            tx_reload.Commit()
        except Exception as e:
            output.print_md("⚠️ Failed to reload family '{}': {}".format(fam_name, str(e)))

        # Tally results and build audit entries per mapping row for this family
        total_families += 1
        family_moved_total = sum(moved_counts.values())
        total_moved += family_moved_total

        for old_subcat, new_subcat in mapping_pairs:
            moved_for_row = moved_counts.get(old_subcat, 0)
            rec = {
                "family": fam_name,
                "old": old_subcat,
                "new": new_subcat,
                "parent": parent_name,
                "moved": moved_for_row,
                "old_deleted": old_deleted_map.get(old_subcat, False),
                "delete_reason": delete_reasons.get(old_subcat, ""),
                "saved_path": save_path
            }
            family_results.append(rec)

        output.print_md("✅ Processed family '{}'. Reassigned elements: {}, saved to: {}".format(
            fam_name, family_moved_total, save_path))

    except Exception:
        output.print_md("❌ Error processing family '{}':\n{}".format(fam_name, traceback.format_exc()))

# --- Project-level purge ---
output.print_md("### 🧹 Project-level purge")
try:
    target_old_names = set([old for _, old, _ in mappings_list])
    subcats = revit.query.get_subcategories(doc=revit.doc, purgable=True)
    to_delete = [sc for sc in subcats if getattr(sc, "Name", None) in target_old_names]

    if to_delete:
        with revit.Transaction("Purge Project Subcategories"):
            revit.delete.delete_elements(to_delete)
        for sc in to_delete:
            name = getattr(sc, "Name", "Unknown")
            parent_name = sc.GraphicsStyleCategory.Name if sc.GraphicsStyleCategory else "Unknown Parent"
            deleted_subcats_project.append((name, parent_name))
            output.print_md("🗑️ Deleted project subcategory '{}' (parent: '{}')".format(name, parent_name))
    else:
        output.print_md("ℹ️ No matching project subcategories found.")
except Exception as e:
    failed_deletes_project.append(("Unknown", "Unknown Parent", str(e)))
    output.print_md("⚠️ Project purge failed:\n{}".format(str(e)))

# --- Audit log ---
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_path = os.path.join(save_folder, "subcategory_cleanup_log_{}.csv".format(timestamp))

try:
    with codecs.open(log_path, "w", "utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        # Header
        w.writerow([
            "Scope", "Family", "Old Subcategory", "New Subcategory",
            "Parent Category", "Elements Moved", "Old Subcategory Deleted",
            "Delete Reason", "Saved Path / Details"
        ])

        # Family-level actions
        for rec in family_results:
            w.writerow([
                "Family",
                rec["family"],
                rec["old"],
                rec["new"],
                rec["parent"],
                rec["moved"],
                "Yes" if rec["old_deleted"] else "No",
                rec["delete_reason"],
                rec["saved_path"]
            ])

        # Project-level deletions
        for name, parent in deleted_subcats_project:
            w.writerow(["Project", "", name, "", parent, "", "Deleted", "", ""])
        for item in failed_deletes_project:
            # Some entries may be tuples of different shapes; handle defensively
            if len(item) == 3:
                name, parent, reason = item
            elif len(item) == 2:
                name, parent = item
                reason = ""
            else:
                name = str(item)
                parent = ""
                reason = ""
            w.writerow(["Project", "", name, "", parent, "", "Failed", reason, ""])
    output.print_md("📝 Audit log written to: {}".format(log_path))
except Exception as e:
    output.print_md("⚠️ Failed to write audit log CSV: {}".format(str(e)))

output.print_md("🎯 Family cleanup, reload, targeted project purge, and audit log complete.")