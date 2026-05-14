# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    GraphicsStyleType,
    LinePattern,
    LinePatternElement,
    LinePatternSegment,
    Transaction,
)
from pyrevit import forms
from System.Collections.Generic import List


uiapp = __revit__
app = uiapp.Application
active_doc = uiapp.ActiveUIDocument.Document


class DocOption(object):
    def __init__(self, doc, is_active=False):
        self.doc = doc
        self.name = doc.Title
        self.path = doc.PathName if doc.PathName else "Unsaved Project"
        self.is_active = is_active
        self.display = "{}{} | {}".format(
            self.name,
            " [active]" if is_active else "",
            self.path,
        )

    def __str__(self):
        return self.display


class LineStyleOption(object):
    def __init__(self, source_doc, category):
        self.source_doc = source_doc
        self.category = category
        self.name = category.Name
        self.display = self._build_display()

    def _build_display(self):
        color = None
        try:
            color = self.category.LineColor
        except Exception:
            color = None

        try:
            proj_weight = self.category.GetLineWeight(GraphicsStyleType.Projection)
        except Exception:
            proj_weight = "?"

        rgb = "?"
        if color is not None:
            rgb = "{},{},{}".format(color.Red, color.Green, color.Blue)

        return "{} | RGB {} | Weight {}".format(self.name, rgb, proj_weight)

    def __str__(self):
        return self.display


def get_open_project_docs(include_read_only=True):
    docs = []
    for doc in app.Documents:
        try:
            if doc.IsFamilyDocument:
                continue
            if (not include_read_only) and doc.IsReadOnly:
                continue
            docs.append(doc)
        except Exception:
            continue
    return docs


def get_line_style_options(source_doc):
    options = []
    lines_cat = source_doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    for subcat in lines_cat.SubCategories:
        try:
            options.append(LineStyleOption(source_doc, subcat))
        except Exception:
            pass
    options.sort(key=lambda x: x.name.lower())
    return options


def find_subcategory_by_name(parent_cat, subcat_name):
    try:
        for subcat in parent_cat.SubCategories:
            if subcat.Name == subcat_name:
                return subcat
    except Exception:
        pass
    return None


def build_line_pattern_lookup(doc):
    lookup = {}
    for pattern in FilteredElementCollector(doc).OfClass(LinePatternElement):
        try:
            line_pattern = pattern.GetLinePattern()
            if line_pattern is not None:
                lookup[line_pattern.Name] = pattern
            else:
                lookup[pattern.Name] = pattern
        except Exception:
            try:
                lookup[pattern.Name] = pattern
            except Exception:
                pass
    return lookup


def ensure_line_pattern_in_target(source_doc, pattern_id, target_doc, target_pattern_lookup, results):
    if pattern_id is None:
        return None

    try:
        if pattern_id.IntegerValue < 0:
            return pattern_id
    except Exception:
        return None

    source_pattern_elem = source_doc.GetElement(pattern_id)
    if source_pattern_elem is None:
        return None

    try:
        source_pattern = source_pattern_elem.GetLinePattern()
    except Exception:
        source_pattern = None

    if source_pattern is None:
        return None

    existing_pattern = target_pattern_lookup.get(source_pattern.Name)
    if existing_pattern is not None:
        return existing_pattern.Id

    try:
        new_pattern = LinePattern(source_pattern.Name)
        segments = List[LinePatternSegment]()
        for seg in source_pattern.GetSegments():
            segments.Add(LinePatternSegment(seg.Type, seg.Length))
        new_pattern.SetSegments(segments)

        created_pattern = LinePatternElement.Create(target_doc, new_pattern)
        target_pattern_lookup[source_pattern.Name] = created_pattern
        if source_pattern.Name not in results["patterns_created"]:
            results["patterns_created"].append(source_pattern.Name)
        return created_pattern.Id
    except Exception as ex:
        results["failed"].append(
            "Line pattern '{}' could not be created: {}".format(source_pattern.Name, ex)
        )
        return None


def copy_style_graphics(source_doc, target_doc, source_subcat, target_subcat, target_pattern_lookup, results):
    try:
        target_subcat.LineColor = source_subcat.LineColor
    except Exception:
        pass

    for gst in [GraphicsStyleType.Projection, GraphicsStyleType.Cut]:
        try:
            line_weight = source_subcat.GetLineWeight(gst)
            if line_weight and line_weight > 0:
                target_subcat.SetLineWeight(line_weight, gst)
        except Exception:
            pass

        try:
            pattern_id = source_subcat.GetLinePatternId(gst)
            target_pattern_id = ensure_line_pattern_in_target(
                source_doc,
                pattern_id,
                target_doc,
                target_pattern_lookup,
                results,
            )
            if target_pattern_id is not None:
                target_subcat.SetLinePatternId(target_pattern_id, gst)
        except Exception:
            pass


def transfer_line_styles(source_doc, target_doc, style_options):
    results = {
        "created": [],
        "updated": [],
        "patterns_created": [],
        "failed": [],
    }

    parent_lines_cat = target_doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    target_pattern_lookup = build_line_pattern_lookup(target_doc)

    tx = Transaction(target_doc, "Transfer Line Styles")
    try:
        tx.Start()
        for style_opt in style_options:
            name = style_opt.name
            source_subcat = style_opt.category

            try:
                target_subcat = find_subcategory_by_name(parent_lines_cat, name)
                if target_subcat is None:
                    target_subcat = target_doc.Settings.Categories.NewSubcategory(parent_lines_cat, name)
                    results["created"].append(name)
                else:
                    results["updated"].append(name)

                copy_style_graphics(
                    source_doc,
                    target_doc,
                    source_subcat,
                    target_subcat,
                    target_pattern_lookup,
                    results,
                )
            except Exception as ex:
                results["failed"].append("{}: {}".format(name, ex))

        tx.Commit()
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        results["failed"].append("Transaction failed: {}".format(ex))

    return results


def main():
    docs = get_open_project_docs(include_read_only=True)
    if len(docs) < 2:
        forms.alert(
            "Open at least two project documents before running this tool.",
            exitscript=True,
        )

    source_options = [DocOption(doc, doc.Equals(active_doc)) for doc in docs]
    source_choice = forms.SelectFromList.show(
        sorted(source_options, key=lambda x: x.display.lower()),
        multiselect=False,
        title="Select Source Project",
        button_name="Next",
        name_attr="display",
    )
    if not source_choice:
        forms.alert("No source project selected.", exitscript=True)

    target_docs = [doc for doc in get_open_project_docs(include_read_only=False) if doc != source_choice.doc]
    if not target_docs:
        forms.alert(
            "No writable target project is available. Open another editable project and try again.",
            exitscript=True,
        )

    target_options = [DocOption(doc) for doc in target_docs]
    target_choice = forms.SelectFromList.show(
        sorted(target_options, key=lambda x: x.display.lower()),
        multiselect=False,
        title="Select Target Project",
        button_name="Next",
        name_attr="display",
    )
    if not target_choice:
        forms.alert("No target project selected.", exitscript=True)

    style_options = get_line_style_options(source_choice.doc)
    if not style_options:
        forms.alert("No line styles were found in the source project.", exitscript=True)

    selected_styles = forms.SelectFromList.show(
        style_options,
        multiselect=True,
        title="Select Line Styles to Transfer",
        button_name="Transfer",
        name_attr="display",
    )
    if not selected_styles:
        forms.alert("No line styles selected.", exitscript=True)

    results = transfer_line_styles(source_choice.doc, target_choice.doc, selected_styles)

    msg_lines = [
        "Source: {}".format(source_choice.doc.Title),
        "Target: {}".format(target_choice.doc.Title),
        "",
        "Styles Created: {}".format(len(results["created"])),
        "Styles Updated: {}".format(len(results["updated"])),
        "Patterns Created: {}".format(len(results["patterns_created"])),
        "Failed: {}".format(len(results["failed"])),
    ]

    if results["patterns_created"]:
        msg_lines.append("")
        msg_lines.append("Transferred line patterns:")
        msg_lines.extend(["- {}".format(name) for name in results["patterns_created"][:20]])
        if len(results["patterns_created"]) > 20:
            msg_lines.append("- ... and {} more".format(len(results["patterns_created"]) - 20))

    if results["failed"]:
        msg_lines.append("")
        msg_lines.append("Errors:")
        msg_lines.extend(["- {}".format(item) for item in results["failed"][:20]])
        if len(results["failed"]) > 20:
            msg_lines.append("- ... and {} more".format(len(results["failed"]) - 20))

    forms.alert("\n".join(msg_lines), title="Transfer Line Styles")


if __name__ == "__main__":
    main()
