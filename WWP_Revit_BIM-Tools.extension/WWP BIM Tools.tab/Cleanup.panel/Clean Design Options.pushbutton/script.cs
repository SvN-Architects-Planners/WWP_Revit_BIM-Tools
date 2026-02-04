using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

public class Script
{
    public static void Execute(UIApplication uiapp)
    {
        UIDocument uidoc = uiapp.ActiveUIDocument;
        if (uidoc == null)
        {
            TaskDialog.Show("WWP BIM Tools", "No active document.");
            return;
        }

        Document doc = uidoc.Document;

        List<ElementId> optionIds = new FilteredElementCollector(doc)
            .OfClass(typeof(DesignOption))
            .WhereElementIsNotElementType()
            .Select(e => e.Id)
            .ToList();

        List<ElementId> setIds = new FilteredElementCollector(doc)
            .OfClass(typeof(DesignOptionSet))
            .WhereElementIsNotElementType()
            .Select(e => e.Id)
            .ToList();

        if (optionIds.Count == 0 && setIds.Count == 0)
        {
            TaskDialog.Show("WWP BIM Tools", "No design options found.");
            return;
        }

        string msg = "Design option sets: " + setIds.Count + "\n" +
                     "Design options: " + optionIds.Count + "\n\n" +
                     "All elements contained in these options will be deleted.\n" +
                     "Continue?";

        TaskDialogResult res = TaskDialog.Show(
            "WWP BIM Tools",
            msg,
            TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

        if (res != TaskDialogResult.Yes)
            return;

        HashSet<ElementId> idsToDelete = new HashSet<ElementId>();
        foreach (ElementId id in optionIds)
            idsToDelete.Add(id);
        foreach (ElementId id in setIds)
            idsToDelete.Add(id);

        using (Transaction t = new Transaction(doc, "Clean Design Options"))
        {
            t.Start();
            doc.Delete(idsToDelete.ToList());
            t.Commit();
        }
    }
}
