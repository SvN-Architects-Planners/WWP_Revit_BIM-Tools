#! python3
from pyrevit import revit, DB, UI, script
from System.Collections.Generic import List

TITLE = "WWP BIM Tools"

doc = revit.doc
if not doc:
    UI.TaskDialog.Show(TITLE, "No active document.")
    script.exit()


def is_ungrouped(elem):
    return elem is not None and elem.GroupId == DB.ElementId.InvalidElementId


def is_symbol_like(elem):
    cat = elem.Category if elem else None
    if not cat:
        return False
    if cat.CategoryType == DB.CategoryType.Annotation:
        return True
    bic = cat.Id.IntegerValue
    return bic in (
        int(DB.BuiltInCategory.OST_DetailComponents),
        int(DB.BuiltInCategory.OST_RepeatingDetail),
    )


def add_ungrouped_ids(elements, ids):
    count = 0
    for elem in elements:
        if is_ungrouped(elem) and elem.Id not in ids:
            ids.add(elem.Id)
            count += 1
    return count


ids_to_delete = set()

line_count = add_ungrouped_ids(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.CurveElement)
    .WhereElementIsNotElementType()
    .ToElements(),
    ids_to_delete,
)

filled_region_count = add_ungrouped_ids(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.FilledRegion)
    .WhereElementIsNotElementType()
    .ToElements(),
    ids_to_delete,
)

tag_count = add_ungrouped_ids(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.IndependentTag)
    .WhereElementIsNotElementType()
    .ToElements(),
    ids_to_delete,
)

area_count = add_ungrouped_ids(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.Area)
    .WhereElementIsNotElementType()
    .ToElements(),
    ids_to_delete,
)

room_count = add_ungrouped_ids(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.Architecture.Room)
    .WhereElementIsNotElementType()
    .ToElements(),
    ids_to_delete,
)

mask_count = add_ungrouped_ids(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.MaskingRegion)
    .WhereElementIsNotElementType()
    .ToElements(),
    ids_to_delete,
)

text_note_count = add_ungrouped_ids(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.TextNote)
    .WhereElementIsNotElementType()
    .ToElements(),
    ids_to_delete,
)

symbols = (
    DB.FilteredElementCollector(doc)
    .OfClass(DB.FamilyInstance)
    .WhereElementIsNotElementType()
    .ToElements()
)

symbol_count = add_ungrouped_ids(
    [e for e in symbols if is_symbol_like(e)],
    ids_to_delete,
)

if not ids_to_delete:
    UI.TaskDialog.Show(TITLE, "No ungrouped elements found for the selected categories.")
    script.exit()

msg = (
    "This will delete ungrouped elements:\n"
    "- Lines: {0}\n"
    "- Filled regions: {1}\n"
    "- Tags: {2}\n"
    "- Areas: {3}\n"
    "- Rooms: {4}\n"
    "- Masking regions: {5}\n"
    "- Text notes: {6}\n"
    "- Symbols/detail items: {7}\n\n"
    "Continue?"
).format(
    line_count,
    filled_region_count,
    tag_count,
    area_count,
    room_count,
    mask_count,
    text_note_count,
    symbol_count,
)

res = UI.TaskDialog.Show(
    TITLE,
    msg,
    UI.TaskDialogCommonButtons.Yes | UI.TaskDialogCommonButtons.No,
)

if res != UI.TaskDialogResult.Yes:
    script.exit()

with DB.Transaction(doc, "Delete Ungrouped Elements") as t:
    t.Start()
    doc.Delete(List[DB.ElementId](list(ids_to_delete)))
    t.Commit()
