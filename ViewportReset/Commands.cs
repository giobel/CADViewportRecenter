using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using System;
using forms = System.Windows.Forms;
using System.Linq;
using Autodesk.AutoCAD.Geometry;
using System.Reflection;

namespace AttributeUpdater
{
    public class Commands
    {
        [CommandMethod("TRISTAN")]
        public void UpdateAttributeInFiles()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;


            // User should input the folder where the dwgs are saved
            PromptResult pr = ed.GetString("\nEnter folder containing DWGs to process: ");

            if (pr.Status != PromptStatus.OK)
                return;

            string pathName = pr.StringResult;

            string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

            // We'll use some counters to keep track
            // of how the processing is going
            int processed = 0, saved = 0, problem = 0;

            var dict = File.ReadLines($"{pathName}\\summary.csv").Select(line => line.Split(',')).ToDictionary(line => line[0], line => line.ToList());

            dict.Remove(dict.Keys.First()); //remove the csv header

            foreach (string fileName in fileNames)
            {
                if (fileName.EndsWith(".dwg", StringComparison.CurrentCultureIgnoreCase))
                {
                    FileInfo fi = new FileInfo(fileName);

                    string name = fi.Name.Substring(0, fi.Name.Length - 4);
                    string outputName = fileName.Substring(0, fileName.Length - 4) + "_updated.dwg";

                    Database db = new Database(false, false);
                    using (db)
                    {
                        try
                        {
                            ed.WriteMessage("\n\nProcessing file: " + fileName);

                            db.ReadDwgFile(fileName, FileShare.ReadWrite, false, "");

                            LayoutManager lm = LayoutManager.Current;

                            lm.CurrentLayout = "Model";

                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                //Attch Xref
                                string PathName = $"{pathName}\\{dict[name][10]}";
                                ObjectId acXrefId = db.AttachXref(PathName, dict[name][10]);

                                if (!acXrefId.IsNull)
                                {
                                    // Attach the DWG reference to the current space
                                    Point3d insPt = new Point3d(0, 0, 0);
                                    using (BlockReference acBlkRef = new BlockReference(insPt, acXrefId))
                                    {
                                        BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                        BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                        modelSpace.AppendEntity(acBlkRef);

                                        trans.AddNewlyCreatedDBObject(acBlkRef, true);
                                    }
                                }

                                lm.CurrentLayout = "Layout1";

                                //forms.MessageBox.Show(lm.CurrentLayout);

                                string currentLo = lm.CurrentLayout;

                                DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                                Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;

                                

                                foreach (ObjectId ID in CurrentLo.GetViewports())
                                {
                                    Viewport VP = trans.GetObject(ID, OpenMode.ForRead) as Viewport;

                                    Point3d revitViewportCentre = new Point3d(double.Parse(dict[name][5]), double.Parse(dict[name][6]), 0);

                                    Point3d revitViewCentreWCS = new Point3d(double.Parse(dict[name][1]), double.Parse(dict[name][2]), 0);

                                    if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < 100)  //Should use the closest viewport, not a fixed distance
                                    {
                                        VP.UpgradeOpen();
                                        double cs = VP.CustomScale; //save the original scale as it changes when we change viewport width and height
                                        VP.CenterPoint = revitViewportCentre; //move the viewport to the revit location
                                        VP.Width = double.Parse(dict[name][8]); //set the width to match the revit width
                                        VP.Height = double.Parse(dict[name][9]); //idem
                                        VP.CustomScale = cs;
                                        //VP.Erase();

                                        TwistViewport(VP.Id, revitViewCentreWCS, DegToRad(double.Parse(dict[name][4])));
                                    }
                                }

                                trans.Commit();
                            }

                            BindXrefs(db);

                            ed.WriteMessage("\nSaving to file: {0}", outputName);

                            db.SaveAs(outputName, DwgVersion.Current);

                            saved++;

                            processed++;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nProblem processing file: {0} - \"{1}\"", fileName, ex.Message);

                            problem++;
                        }
                    }
                }
            }
            ed.WriteMessage(
              "\n\nSuccessfully processed {0} files, of which {1} had " +
              "attributes to update and an additional {2} had errors " +
              "during reading/processing.",
              processed,
              saved,
              problem
            );
        }


        public void BindXrefs(Database db)

        {
            ObjectIdCollection xrefCollection = new ObjectIdCollection();

            using (XrefGraph xg = db.GetHostDwgXrefGraph(false))

            {
                int numOfNodes = xg.NumNodes;

                for (int cnt = 0; cnt < xg.NumNodes; cnt++)

                {
                    XrefGraphNode xNode = xg.GetXrefNode(cnt) as XrefGraphNode;

                    if (!xNode.Database.Filename.Equals(db.Filename))

                    {

                        if (xNode.XrefStatus == XrefStatus.Resolved)

                        {

                            xrefCollection.Add(xNode.BlockTableRecordId);

                        }

                    }
                }
            }

            if (xrefCollection.Count != 0)

                db.BindXrefs(xrefCollection, true);
        }

        public static void zoomExtentTest()

        {
            Object acadObject = Application.AcadApplication;

            acadObject.GetType().InvokeMember("ZoomExtents", BindingFlags.InvokeMethod, null, acadObject, null);
        }

        static void Zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            int nCurVport = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
            // Get the extents of the current space no points
            // or only a center point is provided
            // Check to see if Model space is current
            if (acCurDb.TileMode == true)
            {
                if (pMin.Equals(new Point3d()) == true &&
               pMax.Equals(new Point3d()) == true)
                {
                    pMin = acCurDb.Extmin;
                    pMax = acCurDb.Extmax;
                }
            }
            else
            {
                // Check to see if Paper space is current
                if (nCurVport == 1)
                {
                    // Get the extents of Paper space
                    if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Pextmin;
                        pMax = acCurDb.Pextmax;
                    }
                }
                else
                {
                    // Get the extents of Model space
                    if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Extmin;
                        pMax = acCurDb.Extmax;
                    }
                }
            }
            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Get the current view
                using (ViewTableRecord acView = acDoc.Editor.GetCurrentView())
                {
                    Extents3d eExtents;
                    // Translate WCS coordinates to DCS
                    Matrix3d matWCS2DCS;
                    matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) *
                    matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                    acView.ViewDirection,
                    acView.Target) * matWCS2DCS;
                    // If a center point is specified, define the min and max
                    // point of the extents
                    // for Center and Scale modes
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pMin = new Point3d(pCenter.X - (acView.Width / 2),
                        pCenter.Y - (acView.Height / 2), 0);
                        pMax = new Point3d((acView.Width / 2) + pCenter.X,
                        (acView.Height / 2) + pCenter.Y, 0);
                    }
                    // Create an extents object using a line
                    using (Line acLine = new Line(pMin, pMax))
                    {
                        eExtents = new Extents3d(acLine.Bounds.Value.MinPoint,
                        acLine.Bounds.Value.MaxPoint);
                    }
                    // Calculate the ratio between the width and height of the current view
                    double dViewRatio;
                    dViewRatio = (acView.Width / acView.Height);
                    // Tranform the extents of the view
                    matWCS2DCS = matWCS2DCS.Inverse();
                    eExtents.TransformBy(matWCS2DCS);
                    double dWidth;
                    double dHeight;
                    Point2d pNewCentPt;
                    // Check to see if a center point was provided (Center and Scale modes)
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        dWidth = acView.Width;
                        dHeight = acView.Height;
                        if (dFactor == 0)
                        {
                            pCenter = pCenter.TransformBy(matWCS2DCS);
                        }
                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                    }
                    else // Working in Window, Extents and Limits mode
                    {
                        // Calculate the new width and height of the current view
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;
                        // Get the center of the view
                        pNewCentPt = new Point2d(((eExtents.MaxPoint.X +
                        eExtents.MinPoint.X) * 0.5),
                        ((eExtents.MaxPoint.Y +
                        eExtents.MinPoint.Y) * 0.5));
                    }
                    // Check to see if the new width fits in current window
                    if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;
                    // Resize and scale the view
                    if (dFactor != 0)
                    {
                        acView.Height = dHeight * dFactor;
                        acView.Width = dWidth * dFactor;
                    }
                    // Set the center of the view
                    acView.CenterPoint = pNewCentPt;
                    // Set the current view
                    acDoc.Editor.SetCurrentView(acView);
                }
                // Commit the changes
                acTrans.Commit();
            }
        }
        static double DegToRad(double deg)
        {
            return deg * Math.PI / 180.0;

        }

        private static void TwistViewport(ObjectId vpId, Point3d target, double angle)
        {


            using (Transaction tran = vpId.Database.TransactionManager.StartTransaction())
            {
                Viewport vport = (Viewport)tran.GetObject(vpId, OpenMode.ForWrite);
                vport.Locked = false;
                vport.ViewDirection = Vector3d.ZAxis;
                vport.ViewTarget = target;
                vport.ViewCenter = Point2d.Origin;
                vport.TwistAngle = Math.PI * 2 - angle;
                //vport.Locked = true;

                tran.Commit();
            }
        }

    }
}