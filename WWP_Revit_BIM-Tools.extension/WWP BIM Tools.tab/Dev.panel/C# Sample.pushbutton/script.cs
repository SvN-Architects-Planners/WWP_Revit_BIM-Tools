using Autodesk.Revit.UI;

public class Script
{
    public static void Execute(UIApplication uiapp)
    {
        TaskDialog.Show("WWP BIM Tools", "Hello from C# (pyRevit).");
    }
}
