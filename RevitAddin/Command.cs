#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using winForm = System.Windows.Forms;

#endregion

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            StringBuilder sb = new StringBuilder();
            string outputFile = "summary.csv";

            ProjectLocation pl = doc.ActiveProjectLocation;
            Transform ttr = pl.GetTotalTransform().Inverse;
            ProjectPosition projPosition = doc.ActiveProjectLocation.GetProjectPosition(new XYZ(0, 0, 0)); //rotation??

            IEnumerable<ViewSheet> allSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType().ToElements().Cast<ViewSheet>();

            //List<string> sheetNumbers = new List<string>();

            //sheetNumbers.Add("S1000");
            //sheetNumbers.Add("S1001");
            //sheetNumbers.Add("S1002");
            //sheetNumbers.Add("S1003");
            //sheetNumbers.Add("S1004");
            //sheetNumbers.Add("S1005");

            try
            {

                using (var form = new Form1())
                {
                    //use ShowDialog to show the form as a modal dialog box. 
                    form.ShowDialog();

                    //if the user hits cancel just drop out of macro
                    if (form.DialogResult == winForm.DialogResult.Cancel)
                    {
                        return Result.Cancelled;
                    }

                 string[] sheetNumbers = form.tBoxSheetNumber.Split(' ');

                foreach (string sheetNumber in sheetNumbers)
                {

                    ViewSheet vs = allSheets.Where(x => x.SheetNumber == sheetNumber).First();

                    Viewport vp = doc.GetElement(vs.GetAllViewports().First()) as Viewport;

                    Autodesk.Revit.DB.View vpPlan = doc.GetElement(vp.ViewId) as Autodesk.Revit.DB.View;

                    //Centrepoint of View
                    BoundingBoxXYZ viewBBox = vpPlan.get_BoundingBox(vpPlan);
                    XYZ viewCentre = Helpers.BBoxCenter(viewBBox, doc, vs); // per Project Base Point
                    XYZ viewCentreWCS = ttr.OfPoint(viewCentre);    // per Survey Point

                    //Centrepoint of Viewport on Sheet
                    BoundingBoxXYZ viewPortBBox = vp.get_BoundingBox(vs);
                    XYZ vpCentreOnSheet = Helpers.BBoxCenter(viewPortBBox, doc, vs);        //viewport centre on sheet

                    string xrefName = sheetNumber + "-xref";

                    Helpers.ExportDWG(doc, vs, "Metrolinx Export Settings", sheetNumber);

                    Helpers.ExportDWG(doc, vpPlan, "Metrolinx Export Settings", xrefName);


                    sb.AppendLine(String.Format("{0},{1},{2},{3},{4}",
                                                sheetNumber,
                                                Helpers.PointToString(viewCentreWCS),
                                                projPosition.Angle * 180 / Math.PI,
                                                Helpers.PointToString(vpCentreOnSheet),
                                                xrefName
                                               )
                                               );
                }

                File.WriteAllText(outputFile, sb.ToString());

                TaskDialog.Show("Done", "Sheets have been exported and the csv has been created");
                }
                return Result.Succeeded;
            }
            catch(Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }
            
        }
    }
}
