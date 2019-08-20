using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class FindOversizedViewport : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {

            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Application app = uiapp.Application;
                Document doc = uidoc.Document;

                string outputFile = @"C:\Temp\report.csv";

                StringBuilder sb = new StringBuilder();

                try
                {
                    File.WriteAllText(outputFile,
                      "Sheet Number," +
                      "Viewport Name," +
                      "Viewport Width," +
                      "Viewport Height," +
                      Environment.NewLine
                     );
                }
                catch
                {
                    TaskDialog.Show("Error", "Opertion cancelled. Please close the log file.");
                    return Result.Failed;
                }

                IEnumerable<ViewSheet> allSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
                    .WhereElementIsNotElementType().ToElements().Cast<ViewSheet>();

                //string sheetNumber = "A2001";

                int n = allSheets.Count();
                string s = "{0} of " + n.ToString() + " sheets processed...";
                string caption = "Find oversized views";

                using (ProgressForm pf = new ProgressForm(caption, s, n))
                {

                    foreach (ViewSheet vs in allSheets)
                    {
                        if (pf.abortFlag)
                            break;

                        ICollection<ElementId> viewports = vs.GetAllViewports();

                        foreach (ElementId vpId in viewports)
                        {
                            try
                            {
                                Viewport vp = doc.GetElement(vpId) as Viewport;
                                View planView = doc.GetElement(vp.ViewId) as View;


                                if (planView.ViewType == ViewType.FloorPlan || planView.ViewType == ViewType.EngineeringPlan || planView.ViewType == ViewType.CeilingPlan)
                                {
                                    //XYZ maxPt = vp.GetBoxOutline().MaximumPoint; too slow
                                    //XYZ minPt = vp.GetBoxOutline().MinimumPoint;
                                    BoundingBoxXYZ bbox = vp.get_BoundingBox(vs);

                                    XYZ maxPt = bbox.Max;
                                    XYZ minPt = bbox.Min;

                                    int width = Convert.ToInt32((maxPt.X - minPt.X) * 304.8);
                                    int height = Convert.ToInt32((maxPt.Y - minPt.Y) * 304.8);

                                    int checkWidth = 823;
                                    int checkHeight = 482;

                                    if (width > checkWidth || height > checkHeight)
                                    sb.AppendLine($"{vs.SheetNumber},{vp.Name},{width},{height}");

                                }


                            }
                            catch
                            {
                                sb.AppendLine($"{vs.SheetNumber}, ERROR");
                            }

                        }
                        pf.Increment();
                    }//close foreach
                }
                //ViewSheet vs = allSheets.Where(x => x.SheetNumber == sheetNumber).First();
                
                File.AppendAllText(outputFile, sb.ToString());

                TaskDialog myDialog = new TaskDialog("Summary");
                myDialog.MainIcon = TaskDialogIcon.TaskDialogIconNone;
                myDialog.MainContent = $"Operation completed.";

                myDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, $"Open Log File {outputFile}", "");

                TaskDialogResult res = myDialog.Show();

                if (TaskDialogResult.CommandLink4 == res)
                {
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.StartInfo.FileName = outputFile;
                    process.Start();
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Result", ex.Message);
                return Result.Cancelled;
            }
        }
    }
}
