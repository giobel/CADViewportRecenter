#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.IFC;
using Autodesk.Revit.UI;
using winForm = System.Windows.Forms;

#endregion

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class ExportXrefs : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            //Goal: convert the Viewport centre coordinates on a Sheet to Project Base Point (and then Survey Point) coordinates.
            //Solution: Find the Viewport's view centre point which is in PBP coordinates.
            //Problem: the Vieport centre does not always match the View centre. Annotations, Matchlines and Grids can affect the extent of the viewport.
            //Solution: hide all these categories and find the Viewport centre that matches the View centre. Then find the View centre point in PBP coordinates and translate it
            //by the vector from Viewport original centre and the centre of the Viewport with the categories hidden.
            //https://thebuildingcoder.typepad.com/blog/2018/03/boston-forge-accelerator-and-aligning-plan-views.html

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}",
                        "Sheet Number",
                        "View Centre WCS-X",
                        "View Centre WCS-Y",
                        "View Centre WCS-Z",
                        "Angle to North",
                        "Viewport Centre-X",
                        "Viewport Centre-Y",
                        "Viewport Centre-Z",
                        "Viewport Width",
                        "Viewport Height",
                        "Xref name"
                       )
                       );

            string outputFile = "summary.csv";

            ProjectLocation pl = doc.ActiveProjectLocation;
            Transform ttr = pl.GetTotalTransform().Inverse;
            ProjectPosition projPosition = doc.ActiveProjectLocation.GetProjectPosition(new XYZ(0, 0, 0)); //rotation

            IEnumerable<ViewSheet> allSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType().ToElements().Cast<ViewSheet>();

            IList<string> dWGExportOptions = DWGExportOptions.GetPredefinedSetupNames(doc);

            List<ViewScheduleOption> viewScheduleOptions = Helpers.GetViewScheduleOptions(doc);

            int counter = 0;

            try
            {

                using (var form = new Form1())
                {
                    form.cboxExportSettingsDataSource = dWGExportOptions;
                    //set the form sheets
                    form.cboxSheetDataSource = viewScheduleOptions;
                    //use ShowDialog to show the form as a modal dialog box. 
                    form.ShowDialog();

                    //if the user hits cancel just drop out of macro
                    if (form.DialogResult == winForm.DialogResult.Cancel)
                    {
                        return Result.Cancelled;
                    }

                    string destinationFolder = form.tBoxDestinationFolder;

                    //string[] sheetNumbers = form.tboxSelectedSheets.Split(' ');
                    List<ViewSheet> selectedSheets = form.tboxSelectedSheets;

                    string exportSettings = form.tBoxExportSettings;

                    int n = form.tboxSelectedSheets.Count;
                    string s = "{0} of " + n.ToString() + " plans exported...";
                    string caption = "Export xrefs";

                    string sheetWithoutArchOrEngViewports = "";

                    using (ProgressForm pf = new ProgressForm(caption, s, n))
                    {
                        foreach (ViewSheet vs in selectedSheets)
                        {
                            if (pf.abortFlag)
                                break;


                            //Collect all the viewports on the sheet
                            ICollection<ElementId> viewportIds = vs.GetAllViewports();

                            //Find the viewports that shows a Floor Plan (Architecture) or Structural Plan (Engineering). 
                            Dictionary<Viewport, View> viewportViewDict = new Dictionary<Viewport, View>();

                            //Find the viewplan shown in the viewport
                            //List<View> vpPlan = new List<View>();

                            //if the sheet does not contain FloorPlan, EngineeringPlan or CeilingPlan, do not export it
                            int hasArchOrStrViewports = 0;

                            foreach (ElementId eid in viewportIds)
                            {
                                Viewport vport = doc.GetElement(eid) as Viewport;
                                View planView = doc.GetElement(vport.ViewId) as View;

                                if (planView.ViewType == ViewType.FloorPlan || planView.ViewType == ViewType.EngineeringPlan || planView.ViewType == ViewType.CeilingPlan || planView.ViewType == ViewType.AreaPlan)
                                {
                                    viewportViewDict.Add(vport, planView);
                                    //vpPlan.Add(planView);
                                    hasArchOrStrViewports += 1;
                                }
                            }


                            if (hasArchOrStrViewports != 0)
                            {
                                //Sheet filename
                                string fileName = vs.LookupParameter("CADD File Name").AsString() ?? vs.SheetNumber;

                                XYZ centroidCheck = null;

                                foreach (Viewport vp in viewportViewDict.Keys)
                                {
                                    View vpPlan = viewportViewDict[vp];
                                    //Get the current Viewport Centre for Autocad Viewport
                                    XYZ unchangedVPcenter = vp.GetBoxCenter();
                                    //Set its Z value to 0
                                    XYZ flattenVPcenter = new XYZ(unchangedVPcenter.X, unchangedVPcenter.Y, 0);
                                    //The current Viewport Centre does not match the view center. Hide all the elements in the view and set the annotation crop to the minimum (3mm). Now the 
                                    //Viewport centre will match the View centre. We can then use the unchangedCenter to changedCenter vector to move the view centerpoint to the original
                                    //viewport centre.

                                    //Instead of hiding categories, we can isolate an empty one. Check that this category (OST_Loads) exists in the model or it will throw an error
                                    ICollection<ElementId> categoryToIsolate = new List<ElementId>();
                                    Categories groups = doc.Settings.Categories;
                                    categoryToIsolate.Add(groups.get_Item(BuiltInCategory.OST_Loads).Id);

                                    //This is the new Viewport centre, aligned with the View centre
                                    XYZ changedVPcenter = null;

                                    // This is the View centre
                                    XYZ centroid = null;

                                    double scale = 304.8;

                                    using (Transaction t = new Transaction(doc, "Hide categories"))
                                    {
                                        t.Start();

                                        vpPlan.IsolateCategoriesTemporary(categoryToIsolate);

                                        //Use the annotation crop region to find the view centroid
                                        ViewCropRegionShapeManager vcr = vpPlan.GetCropRegionShapeManager();
                                        //Set the annotation offset to the minimum (3mm)
                                        vcr.BottomAnnotationCropOffset = 3 / scale;
                                        vcr.TopAnnotationCropOffset = 3 / scale;
                                        vcr.LeftAnnotationCropOffset = 3 / scale;
                                        vcr.RightAnnotationCropOffset = 3 / scale;
                                        //Get the Viewport Center. This will match the View centroid
                                        changedVPcenter = vp.GetBoxCenter();

                                        //Find the view centroid using the annotation crop shape (it should always be a rectangle, while the cropbox shape can be a polygon).
                                        CurveLoop cloop = vcr.GetAnnotationCropShape();
                                        List<XYZ> pts = new List<XYZ>();

                                        foreach (Curve crv in cloop)
                                        {
                                            pts.Add(crv.GetEndPoint(0));
                                            pts.Add(crv.GetEndPoint(1));
                                        }

                                        //View centroid with elements hidden
                                        centroid = Helpers.GetCentroid(pts, pts.Count);

                                        t.RollBack();
                                    }

                                    //Set ChangedVP center Z value to 0
                                    XYZ flattenChangedVPcenter = new XYZ(changedVPcenter.X, changedVPcenter.Y, 0);

                                    //This is the vector from the Viewport original centre to Viewport centre with all the elements hidden and the cropbox set to 3mm evenly
                                    XYZ viewPointsVector = (flattenVPcenter - flattenChangedVPcenter) * vpPlan.Scale;

                                    //View center adjusted to Viewport original centre (the one to be used in Autocad)
                                    XYZ translatedCentroid = centroid + viewPointsVector;

                                    //View center per Survey Point coordinates
                                    XYZ viewCentreWCS = ttr.OfPoint(translatedCentroid);

                                    XYZ viewCentreWCSZ = new XYZ(viewCentreWCS.X, viewCentreWCS.Y, vpPlan.get_BoundingBox(vpPlan).Transform.Origin.Z);

                                    //Viewport outline width and height to be used to update the autocad viewport
                                    XYZ maxPt = vp.GetBoxOutline().MaximumPoint;
                                    XYZ minPt = vp.GetBoxOutline().MinimumPoint;
                                    int width = Convert.ToInt32((maxPt.X - minPt.X) * 304.8);
                                    int height = Convert.ToInt32((maxPt.Y - minPt.Y) * 304.8);


                                    //Suffix to xref
                                    string xrefName = $"{vs.SheetNumber}-{vp.Id}-xref";


                                    if (!Helpers.ExportDWG(doc, vpPlan, exportSettings, xrefName, destinationFolder))
                                    {
                                        TaskDialog.Show("Error", "Check that the destination folder exists");
                                    }
                                    else
                                    {
                                        counter += 1;
                                    }

                                    sb.AppendLine(String.Format("{0},{1},{2},{3},{4},{5},{6}",
                                                    fileName,
                                                    Helpers.PointToString(viewCentreWCSZ),
                                                    projPosition.Angle * 180 / Math.PI,
                                                    Helpers.PointToString(flattenVPcenter),
                                                    width.ToString(),
                                                    height.ToString(),
                                                    xrefName
                                                   )
                                                   );

                                }//close foreach

                            }
                            else
                            {
                                sheetWithoutArchOrEngViewports += $"{vs.SheetNumber}\n";
                            }



                            pf.Increment();
                        }
                    }
                    File.WriteAllText($"{destinationFolder}\\{outputFile}", sb.ToString());

                    TaskDialog.Show("Done", $"{counter} plans have been exported and the csv has been created.\nNot exported:\n{sheetWithoutArchOrEngViewports}");
                }//close form
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

        }
    }
}
