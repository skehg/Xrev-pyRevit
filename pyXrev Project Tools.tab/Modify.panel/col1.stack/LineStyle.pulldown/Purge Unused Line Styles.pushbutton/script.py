# -*- coding: utf-8 -*-
"""
Purge Unused Line Styles
------------------------
Finds line styles in the active Revit project that are not assigned to any
curve elements and allows the user to delete only the styles selected in the
pick list.
"""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    CurveElement,
    FilteredElementCollector,
    GraphicsStyleType,
    Transaction,
)
from pyrevit import forms, revit, script


doc = revit.doc
output = script.get_output()


class LineStyleOption(object):
    def __init__(self, category, usage_count, blocked_reason=None):
        self.category = category
        self.name = category.Name
        self.usage_count = usage_count
        self.blocked_reason = blocked_reason
        self.display = self._build_display()

    def _build_display(self):
        rgb = "?, ?, ?"
        weight = "?"

        try:
            color = self.category.LineColor
            if color is not None:
                rgb = "{}, {}, {}".format(color.Red, color.Green, color.Blue)
        except Exception:
            pass

        try:
            weight = self.category.GetLineWeight(GraphicsStyleType.Projection)
        except Exception:
            pass

        return "{} | Weight {} | RGB {}".format(self.name, weight, rgb)

    def __str__(self):
        return self.display


def get_lines_parent_category(revit_doc):
    try:
        return revit_doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    except Exception:
        return None


def get_used_style_name_counts(revit_doc):
    counts = {}
    collector = FilteredElementCollector(revit_doc).OfClass(CurveElement).WhereElementIsNotElementType()

    for curve in collector:
        style_name = None

        try:
            line_style = curve.LineStyle
            if line_style is not None:
                style_name = getattr(line_style, "Name", None)
                try:
                    style_cat = getattr(line_style, "GraphicsStyleCategory", None)
                    if style_cat is not None and getattr(style_cat, "Name", None):
                        style_name = style_cat.Name
                except Exception:
                    pass
        except Exception:
            continue

        if style_name:
            counts[style_name] = counts.get(style_name, 0) + 1

    return counts


def is_user_created_line_style(category):
    if category is None:
        return False

    try:
        if category.Id.IntegerValue < 0:
            return False
    except Exception:
        return False

    try:
        parent = category.Parent
        if parent is None:
            return False
    except Exception:
        return False

    return True


def find_unused_line_styles(revit_doc):
    lines_cat = get_lines_parent_category(revit_doc)
    if lines_cat is None:
        return [], {}, 0

    usage_counts = get_used_style_name_counts(revit_doc)
    purgeable = []
    ignored_count = 0

    for subcat in lines_cat.SubCategories:
        try:
            if not is_user_created_line_style(subcat):
                ignored_count += 1
                continue

            style_name = subcat.Name
            use_count = usage_counts.get(style_name, 0)
            if use_count > 0:
                continue

            purgeable.append(LineStyleOption(subcat, use_count))
        except Exception:
            pass

    purgeable.sort(key=lambda x: x.name.lower())
    return purgeable, usage_counts, ignored_count


def delete_line_styles(revit_doc, selected_options):
    deleted = []
    failed = []

    tx = Transaction(revit_doc, "Purge Unused Line Styles")
    try:
        tx.Start()
        for option in selected_options:
            if not is_user_created_line_style(option.category):
                continue
            try:
                revit_doc.Delete(option.category.Id)
                deleted.append(option.name)
            except Exception as ex:
                failed.append("{}: {}".format(option.name, ex))

        tx.Commit()
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        failed.append("Transaction failed: {}".format(ex))

    return deleted, failed


def print_summary(purgeable, deleted, failed):
    output.print_md("## Purge Unused Line Styles")
    output.print_md("- Candidate styles found: {}".format(len(purgeable)))
    output.print_md("- Deleted styles: {}".format(len(deleted)))
    output.print_md("- Failures: {}".format(len(failed)))

    if deleted:
        output.print_md("### Deleted")
        for name in deleted:
            output.print_md("- {}".format(name))

    if failed:
        output.print_md("### Failures")
        for item in failed:
            output.print_md("- {}".format(item))


def main():
    if doc is None:
        forms.alert("No active Revit document was found.", exitscript=True)

    if doc.IsFamilyDocument:
        forms.alert(
            "Run this tool from a project document. Family documents do not support project line-style purging.",
            exitscript=True,
        )

    purgeable, usage_counts, ignored_count = find_unused_line_styles(doc)

    if not purgeable:
        forms.alert("No unused user-created line styles were found in the active project.", exitscript=True)

    selected = forms.SelectFromList.show(
        purgeable,
        multiselect=True,
        title="Purge Unused Line Styles",
        button_name="Delete Selected",
        name_attr="display",
    )

    if not selected:
        forms.alert("No line styles were selected.", exitscript=True)

    confirm = forms.alert(
        "Delete {} unused line style(s) from the current project?".format(len(selected)),
        yes=True,
        no=True,
    )
    if not confirm:
        script.exit()

    deleted, failed = delete_line_styles(doc, selected)
    print_summary(purgeable, deleted, failed)

    message = "Deleted {} selected line style(s).".format(len(deleted))
    if failed:
        message += "\n{} selected style(s) could not be removed. See the pyRevit output window for details.".format(len(failed))

    forms.alert(message)


if __name__ == "__main__":
    main()
