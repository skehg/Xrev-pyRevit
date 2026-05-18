# pylint: disable=E0401,W0613,C0103,C0111
# -*- coding: utf-8 -*-
import sys
from Autodesk.Revit.DB import Transaction, TransactionGroup
from pyrevit import revit, DB, script, forms
from pyrevit.framework import List

# ----------------------------------------------------------------
# Step 1 – pick destination documents
# ----------------------------------------------------------------
open_docs = forms.select_open_docs(title="Select Destination Documents")
if not open_docs:
    sys.exit(0)

# ----------------------------------------------------------------
# Step 2 – pick source drafting views from the current document
# ----------------------------------------------------------------
src_views = forms.select_views(
    title="Select Drafting Views to Copy",
    filterfunc=lambda x: x.ViewType == DB.ViewType.DraftingView,
    use_selection=True,
)

if not src_views:
    forms.alert("No Drafting Views selected.")
    sys.exit(0)


# ----------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------
class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    """Keep destination types when name clashes occur during copy."""
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


def get_drafting_vft(dest_doc):
    """Return the first ViewFamilyType for Drafting views in dest_doc."""
    vfts = DB.FilteredElementCollector(dest_doc) \
              .OfClass(DB.ViewFamilyType) \
              .ToElements()
    for vft in vfts:
        if vft.ViewFamily == DB.ViewFamily.Drafting:
            return vft
    return None


# ----------------------------------------------------------------
# Step 3 – copy
# ----------------------------------------------------------------
logger = script.get_logger()

total_operations = len(src_views) * len(open_docs)
current_operation = 0
skipped_docs = []

with forms.ProgressBar(cancellable=True) as pb:
    for dest_doc in open_docs:

        pb.title = "Processing Document: {}".format(dest_doc.Title)
        pb.update_progress(current_operation, total_operations)

        drafting_vft = get_drafting_vft(dest_doc)
        if not drafting_vft:
            forms.alert(
                "Could not find a Drafting View Family Type in target document.\n"
                "Skipping document: {}".format(dest_doc.Title)
            )
            skipped_docs.append(dest_doc.Title)
            current_operation += len(src_views)
            continue

        # Collect existing drafting view names in destination
        all_dest_views = revit.query.get_all_views(doc=dest_doc)
        existing_names = [
            revit.query.get_name(v)
            for v in all_dest_views
            if v.ViewType == DB.ViewType.DraftingView
        ]

        tg = TransactionGroup(dest_doc, "Copy Drafting Views to document")
        tg.Start()

        for src_view in src_views:
            if pb.cancelled:
                tg.RollBack()
                forms.alert("Operation cancelled.")
                sys.exit(0)

            view_name = revit.query.get_name(src_view)
            pb.title = "Processing: {} > {}".format(dest_doc.Title, view_name)
            pb.update_progress(current_operation, total_operations)

            # Collect elements in the source drafting view
            view_elements = DB.FilteredElementCollector(
                revit.doc, src_view.Id
            ).ToElements()

            elements_to_copy = []
            for el in view_elements:
                if isinstance(el, DB.Element) and el.Category:
                    elements_to_copy.append(el.Id)

            t = Transaction(dest_doc, "Copy Drafting View: {}".format(view_name))
            t.Start()

            # Create a fresh drafting view in the destination document
            new_view = DB.ViewDrafting.Create(dest_doc, drafting_vft.Id)

            if elements_to_copy:
                options = DB.CopyPasteOptions()
                options.SetDuplicateTypeNamesHandler(CopyUseDestination())
                try:
                    copied_ids = DB.ElementTransformUtils.CopyElements(
                        src_view,
                        List[DB.ElementId](elements_to_copy),
                        new_view,
                        None,
                        options,
                    )

                    # Replicate per-element graphic overrides
                    for dest_id, src_id in zip(copied_ids, elements_to_copy):
                        try:
                            new_view.SetElementOverrides(
                                dest_id, src_view.GetElementOverrides(src_id)
                            )
                        except Exception as ex:
                            logger.warning(
                                "Could not copy overrides for element {}: {}".format(
                                    src_id, ex
                                )
                            )
                except Exception as ex:
                    logger.warning(
                        "Error copying elements from '{}': {}".format(view_name, ex)
                    )

            # Resolve duplicate name
            new_name = view_name
            counter = 0
            while new_name in existing_names:
                counter += 1
                new_name = "{} (Duplicate {})".format(view_name, counter)

            revit.update.set_name(new_view, new_name)
            new_view.Scale = src_view.Scale
            existing_names.append(new_name)

            t.Commit()
            current_operation += 1

        tg.Assimilate()

# ----------------------------------------------------------------
# Step 4 – summary alert
# ----------------------------------------------------------------
processed_docs = [d for d in open_docs if d.Title not in skipped_docs]
if processed_docs:
    view_count = len(src_views)
    doc_count = len(processed_docs)
    view_text = "drafting view" + ("s" if view_count > 1 else "")

    details = []
    details.append("COPIED {}:".format(view_text.upper()))
    for v in src_views:
        details.append(u"\u2022 {}".format(revit.query.get_name(v)))
    details.append("\nTO DOCUMENT{}:".format("S" if doc_count > 1 else ""))
    for d in processed_docs:
        details.append(u"\u2022 {}".format(d.Title))
    if skipped_docs:
        details.append("\nSKIPPED DOCUMENTS:")
        for name in skipped_docs:
            details.append(u"\u2022 {}".format(name))

    forms.alert(
        "{} {} copied to {} document{}.".format(
            view_count, view_text, doc_count, "s" if doc_count > 1 else ""
        ),
        expanded="\n".join(details),
    )
