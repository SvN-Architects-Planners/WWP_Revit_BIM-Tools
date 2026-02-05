#! python3
from pyrevit import revit, DB, UI, script
from System.Collections.Generic import List

TITLE = "WWP BIM Tools"

doc = revit.doc
if not doc:
    UI.TaskDialog.Show(TITLE, "No active document.")
    script.exit()


def is_detail_group(cat):
    return cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_IOSDetailGroups)


def is_model_group(cat):
    return cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_IOSModelGroups)


detail_curve_ids = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.DetailCurve)
    .WhereElementIsNotElementType()
    .ToElementIds()
)

group_instances = (
    DB.FilteredElementCollector(doc)
    .OfClass(DB.Group)
    .WhereElementIsNotElementType()
    .ToElements()
)

group_types = (
    DB.FilteredElementCollector(doc)
    .OfClass(DB.GroupType)
    .WhereElementIsElementType()
    .ToElements()
)

detail_group_instance_count = len([g for g in group_instances if is_detail_group(g.Category)])
model_group_instance_count = len([g for g in group_instances if is_model_group(g.Category)])
detail_group_type_count = len([g for g in group_types if is_detail_group(g.Category)])
model_group_type_count = len([g for g in group_types if is_model_group(g.Category)])

ids_to_delete = set(detail_curve_ids)

for g in group_instances:
    if is_detail_group(g.Category) or is_model_group(g.Category):
        ids_to_delete.add(g.Id)

for gt in group_types:
    if is_detail_group(gt.Category) or is_model_group(gt.Category):
        ids_to_delete.add(gt.Id)

if not ids_to_delete:
    UI.TaskDialog.Show(TITLE, "No detail lines or groups found.")
    script.exit()

msg = (
    "This will delete:\n"
    "- Detail lines: {0}\n"
    "- Detail group instances: {1}\n"
    "- Model group instances: {2}\n"
    "- Detail group types: {3}\n"
    "- Model group types: {4}\n\n"
    "Continue?"
).format(
    len(detail_curve_ids),
    detail_group_instance_count,
    model_group_instance_count,
    detail_group_type_count,
    model_group_type_count,
)

res = UI.TaskDialog.Show(
    TITLE,
    msg,
    UI.TaskDialogCommonButtons.Yes | UI.TaskDialogCommonButtons.No,
)

if res != UI.TaskDialogResult.Yes:
    script.exit()

with DB.Transaction(doc, "Delete Detail Lines and Groups") as t:
    t.Start()
    doc.Delete(List[DB.ElementId](list(ids_to_delete)))
    t.Commit()
