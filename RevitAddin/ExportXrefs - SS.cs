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
    public class ExportXrefsSupersed : IExternalCommand
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
            ProjectPosition projPosition = doc.ActiveProjectLocation.GetProjectPosition(new XYZ(0, 0, 0)); //rotation??

            IEnumerable<ViewSheet> allSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType().ToElements().Cast<ViewSheet>();

            int counter = 0;

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

                    string destinationFolder = form.tBoxDestinationFolder;

                    string[] sheetNumbers = form.tBoxSheetNumber.Split(' ');

                    string exportSettings = form.tBoxExportSettings;

                    foreach (string sheetNumber in sheetNumbers)
                    {

                        ViewSheet vs = allSheets.Where(x => x.SheetNumber == sheetNumber).First();

                        //Viewport vp = doc.GetElement(vs.GetAllViewports().Where(x => x.GetType().Name == "Floor Plan").First()) as Viewport;

                        ICollection<ElementId> viewportIds = vs.GetAllViewports();

                        Viewport vp = null;
                        View vpPlan = null;

                        foreach (ElementId eid in viewportIds)
                        {
                            Viewport vport = doc.GetElement(eid) as Viewport;
                            View planView = doc.GetElement(vport.ViewId) as View;

                            if (planView.ViewType== ViewType.FloorPlan || planView.ViewType==ViewType.EngineeringPlan)
                            {
                                vp = vport;
                                vpPlan = planView;
                            }
                        }

                        ViewCropRegionShapeManager vcr = vpPlan.GetCropRegionShapeManager();

                        IList<CurveLoop> cloop = vcr.GetCropShape(); //view crop outline

                        CurveLoop annotationloop = vcr.GetAnnotationCropShape();

                        double area = ExporterIFCUtils.ComputeAreaOfCurveLoops(cloop);


                        double scale = vpPlan.Scale;
                        double left = vcr.LeftAnnotationCropOffset * scale;
                        double right = vcr.RightAnnotationCropOffset * scale;
                        double top = vcr.TopAnnotationCropOffset *scale;
                        double bottom = vcr.BottomAnnotationCropOffset *scale;


                        List<Curve> crvs = new List<Curve>();

                        List<XYZ> pts = new List<XYZ>();

                        ////Option 1 centroid of view crop
                        //foreach (Curve crv in cloop.First())
                        //{

                        //    crvs.Add(crv);
                        //    pts.Add(crv.GetEndPoint(0));
                        //    pts.Add(crv.GetEndPoint(1));

                        //    //Element e = doc.Create.NewDetailCurve(vpPlan, crv);
                        //}


                        //Option 2 centroid of annotation crop
                        foreach (Curve crv in annotationloop)
                        {

                            pts.Add(crv.GetEndPoint(0));
                            pts.Add(crv.GetEndPoint(1));

                            //Element e = doc.Create.NewDetailCurve(vpPlan, crv);
                        }

                       
                        XYZ centroid = Helpers.GetCentroid(pts, pts.Count);
                        //XYZ viewCentreWCS = ttr.OfPoint(centroid);    // per Survey Point

                        //Option 1 - Centroid of CropShape + Annotation Crop (Annotation crop matching Viewport boundary)
                        //XYZ centroidAnnotationCrop = new XYZ(centroid.X + (-left + right) / 2, centroid.Y + (top - bottom) / 2, 0);
                        //XYZ viewCentreWCS = ttr.OfPoint(centroidAnnotationCrop);    // per Survey Point

                        //Option 2 - Centroid of Annotation Crop Shape
                        XYZ viewCentreWCS = ttr.OfPoint(centroid);    // per Survey Point

                        //Viewport outline width and height to be used to update the autocad viewport
                        XYZ maxPt = vp.GetBoxOutline().MaximumPoint;
                        XYZ minPt = vp.GetBoxOutline().MinimumPoint;

                        int width = Convert.ToInt32((maxPt.X - minPt.X) * 304.8);
                        int height = Convert.ToInt32((maxPt.Y - minPt.Y) * 304.8);

                        ////Centrepoint of View
                        //BoundingBoxXYZ viewBBox = vpPlan.get_BoundingBox(vpPlan);
                        //XYZ viewCentre = Helpers.BBoxCenter(viewBBox, doc, vs); // per Project Base Point
                        //XYZ viewCentreWCS = ttr.OfPoint(viewCentre);    // per Survey Point

                        //Centrepoint of Viewport on Sheet
                        BoundingBoxXYZ viewPortBBox = vp.get_BoundingBox(vs);
                        //XYZ vpCentreOnSheet = Helpers.BBoxCenter(viewPortBBox, doc, vs);        //viewport centre on sheet

                        XYZ vpCentreOnSheet = vp.GetBoxCenter();
                        //XYZ vpCentreOnSheet = new XYZ(originalVpCentreOnSheet.X + (left - right)/2, originalVpCentreOnSheet.Y + (-top + bottom)/2, 0);


                        string xrefName = sheetNumber + "-xref";

                        

                        if (!Helpers.ExportDWG(doc, vpPlan, exportSettings, xrefName, destinationFolder))
                        {
                            TaskDialog.Show("Error", "Check that the destination folder exists or the Export Settings exists");
                            
                        }
                        else
                        {
                            counter += 1;
                        }

                        sb.AppendLine(String.Format("{0},{1},{2},{3},{4},{5},{6}",
                                                sheetNumber,
                                                Helpers.PointToString(viewCentreWCS),
                                                projPosition.Angle * 180 / Math.PI,
                                                Helpers.PointToString(vpCentreOnSheet),
                                                width.ToString(),
                                                height.ToString(),
                                                xrefName
                                               )
                                               );
                    }

                    File.WriteAllText(outputFile, sb.ToString());

                    TaskDialog.Show("Done", $"{counter} plans have been exported and the csv has been created");
                }
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
