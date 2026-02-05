#! python3
from pyrevit import revit, DB, UI, script
from System.Collections.Generic import List

TITLE = "WWP BIM Tools"

doc = revit.doc
if not doc:
    UI.TaskDialog.Show(TITLE, "No active document.")
    script.exit()

option_ids = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.DesignOption)
    .WhereElementIsNotElementType()
    .ToElementIds()
)

set_ids = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.DesignOptionSet)
    .WhereElementIsNotElementType()
    .ToElementIds()
)

if not option_ids and not set_ids:
    UI.TaskDialog.Show(TITLE, "No design options found.")
    script.exit()

msg = (
    "Design option sets: {0}\n"
    "Design options: {1}\n\n"
    "All elements contained in these options will be deleted.\n"
    "Continue?"
).format(len(set_ids), len(option_ids))

res = UI.TaskDialog.Show(
    TITLE,
    msg,
    UI.TaskDialogCommonButtons.Yes | UI.TaskDialogCommonButtons.No,
)

if res != UI.TaskDialogResult.Yes:
    script.exit()

ids_to_delete = set(option_ids + set_ids)

with DB.Transaction(doc, "Clean Design Options") as t:
    t.Start()
    doc.Delete(List[DB.ElementId](list(ids_to_delete)))
    t.Commit()
