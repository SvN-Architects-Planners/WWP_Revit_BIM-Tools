using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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

        HashSet<ElementId> idsToDelete = new HashSet<ElementId>();

        int lineCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(CurveElement))
                .WhereElementIsNotElementType(),
            idsToDelete);

        int filledRegionCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegion))
                .WhereElementIsNotElementType(),
            idsToDelete);

        int tagCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(IndependentTag))
                .WhereElementIsNotElementType(),
            idsToDelete);

        int areaCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(Area))
                .WhereElementIsNotElementType(),
            idsToDelete);

        int roomCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(Room))
                .WhereElementIsNotElementType(),
            idsToDelete);

        int maskCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(MaskingRegion))
                .WhereElementIsNotElementType(),
            idsToDelete);

        int textNoteCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(TextNote))
                .WhereElementIsNotElementType(),
            idsToDelete);

        int symbolCount = AddUngroupedIds(
            new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .Where(IsSymbolLike),
            idsToDelete);

        if (idsToDelete.Count == 0)
        {
            TaskDialog.Show("WWP BIM Tools", "No ungrouped elements found for the selected categories.");
            return;
        }

        string msg = "This will delete ungrouped elements:\n" +
                     "- Lines: " + lineCount + "\n" +
                     "- Filled regions: " + filledRegionCount + "\n" +
                     "- Tags: " + tagCount + "\n" +
                     "- Areas: " + areaCount + "\n" +
                     "- Rooms: " + roomCount + "\n" +
                     "- Masking regions: " + maskCount + "\n" +
                     "- Text notes: " + textNoteCount + "\n" +
                     "- Symbols/detail items: " + symbolCount + "\n\n" +
                     "Continue?";

        TaskDialogResult res = TaskDialog.Show(
            "WWP BIM Tools",
            msg,
            TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

        if (res != TaskDialogResult.Yes)
            return;

        using (Transaction t = new Transaction(doc, "Delete Ungrouped Elements"))
        {
            t.Start();
            doc.Delete(idsToDelete.ToList());
            t.Commit();
        }
    }

    private static int AddUngroupedIds(IEnumerable<Element> elements, HashSet<ElementId> ids)
    {
        int count = 0;
        foreach (Element e in elements)
        {
            if (IsUngrouped(e) && ids.Add(e.Id))
                count++;
        }
        return count;
    }

    private static bool IsUngrouped(Element e)
    {
        return e != null && e.GroupId == ElementId.InvalidElementId;
    }

    private static bool IsSymbolLike(Element e)
    {
        Category cat = e != null ? e.Category : null;
        if (cat == null)
            return false;

        if (cat.CategoryType == CategoryType.Annotation)
            return true;

        int bic = cat.Id.IntegerValue;
        return bic == (int)BuiltInCategory.OST_DetailComponents
            || bic == (int)BuiltInCategory.OST_RepeatingDetail;
    }
}
