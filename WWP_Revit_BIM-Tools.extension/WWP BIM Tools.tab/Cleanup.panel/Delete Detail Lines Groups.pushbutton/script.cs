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

        List<ElementId> detailCurveIds = new FilteredElementCollector(doc)
            .OfClass(typeof(DetailCurve))
            .WhereElementIsNotElementType()
            .Select(e => e.Id)
            .ToList();

        List<Group> groupInstances = new FilteredElementCollector(doc)
            .OfClass(typeof(Group))
            .WhereElementIsNotElementType()
            .Cast<Group>()
            .ToList();

        List<GroupType> groupTypes = new FilteredElementCollector(doc)
            .OfClass(typeof(GroupType))
            .WhereElementIsElementType()
            .Cast<GroupType>()
            .ToList();

        int detailGroupInstanceCount = groupInstances.Count(g => IsDetailGroup(g.Category));
        int modelGroupInstanceCount = groupInstances.Count(g => IsModelGroup(g.Category));
        int detailGroupTypeCount = groupTypes.Count(gt => IsDetailGroup(gt.Category));
        int modelGroupTypeCount = groupTypes.Count(gt => IsModelGroup(gt.Category));

        HashSet<ElementId> idsToDelete = new HashSet<ElementId>();

        foreach (ElementId id in detailCurveIds)
            idsToDelete.Add(id);

        foreach (Group g in groupInstances)
        {
            if (IsDetailGroup(g.Category) || IsModelGroup(g.Category))
                idsToDelete.Add(g.Id);
        }

        foreach (GroupType gt in groupTypes)
        {
            if (IsDetailGroup(gt.Category) || IsModelGroup(gt.Category))
                idsToDelete.Add(gt.Id);
        }

        if (idsToDelete.Count == 0)
        {
            TaskDialog.Show("WWP BIM Tools", "No detail lines or groups found.");
            return;
        }

        string msg = "This will delete:\n" +
                     "- Detail lines: " + detailCurveIds.Count + "\n" +
                     "- Detail group instances: " + detailGroupInstanceCount + "\n" +
                     "- Model group instances: " + modelGroupInstanceCount + "\n" +
                     "- Detail group types: " + detailGroupTypeCount + "\n" +
                     "- Model group types: " + modelGroupTypeCount + "\n\n" +
                     "Continue?";

        TaskDialogResult res = TaskDialog.Show(
            "WWP BIM Tools",
            msg,
            TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

        if (res != TaskDialogResult.Yes)
            return;

        using (Transaction t = new Transaction(doc, "Delete Detail Lines and Groups"))
        {
            t.Start();
            doc.Delete(idsToDelete.ToList());
            t.Commit();
        }
    }

    private static bool IsDetailGroup(Category cat)
    {
        return cat != null && cat.Id.IntegerValue == (int)BuiltInCategory.OST_IOSDetailGroups;
    }

    private static bool IsModelGroup(Category cat)
    {
        return cat != null && cat.Id.IntegerValue == (int)BuiltInCategory.OST_IOSModelGroups;
    }
}
