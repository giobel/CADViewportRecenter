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
using Autodesk.AutoCAD.PlottingServices;
using System.Collections.Generic;

namespace AttributeUpdater
{
    public class Commands
    {
        [CommandMethod("MERGEANDBIND")]
        public void MergeAndBind()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2017/ENU/AutoCAD-NET/files/GUID-FAC1A5EB-2D9E-497B-8FD9-E11D2FF87B93-htm.html

            //https://adndevblog.typepad.com/autocad/2012/07/using-readdwgfile-with-net-attachxref-or-objectarx-acdbattachxref.html
            //Database oldDb = HostApplicationServices.WorkingDatabase; //is it necessary?

            

            // User should input the folder where the dwgs are saved
            PromptResult pr = ed.GetString("\nEnter folder containing DWGs to process: ");

            if (pr.Status != PromptStatus.OK)
                return;

            string pathName = pr.StringResult;

            string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

            // We'll use some counters to keep track
            // of how the processing is going
            int processed = 0, saved = 0, problem = 0;

            //var dict = File.ReadLines($"{pathName}\\summary.csv").Select(line => line.Split(',')).ToDictionary(line => line[0], line => line.ToList());

            //dict.Remove(dict.Keys.First()); //remove the csv header

            //using a Sheet Object
            var logFile = File.ReadAllLines($"{pathName}\\summary.csv").Select(line => line.Split(',')).ToList<string[]>();
            logFile.RemoveAt(0);

            List<SheetObject> sheetsList = new List<SheetObject>();

            foreach (string[] item in logFile)
            {
                XYZ vc = new XYZ(Convert.ToDouble(item[1]), Convert.ToDouble(item[2]), Convert.ToDouble(item[3]));
                XYZ vpCentre = new XYZ(Convert.ToDouble(item[5]), Convert.ToDouble(item[6]), Convert.ToDouble(item[7]));

                sheetsList.Add(new SheetObject(item[0], vc, Convert.ToDouble(item[4]), vpCentre, Convert.ToDouble(item[8]), Convert.ToDouble(item[9]), item[10]));
            }

            //foreach (string fileName in dict.Keys)
            foreach (SheetObject sheetObject in sheetsList)
            {

                string name = sheetObject.sheetName;
                string filePath = $"{pathName}\\{sheetObject.sheetName}.dwg";
                string outputPath = $"{pathName}\\{sheetObject.sheetName}.dwg";


                //Database db = new Database(false, false);
                Database db = new Database(false, true);
                using (db)
                {
                    try
                    {
                        ed.WriteMessage($"\n\nProcessing file: {filePath} ");

                        db.ReadDwgFile(filePath, FileShare.ReadWrite, true, "");
                        db.CloseInput(true);

                        LayoutManager lm = LayoutManager.Current;

                        lm.CurrentLayout = "Model"; //is it necessary?

                        using (Transaction trans = db.TransactionManager.StartTransaction())
                        {
                            //Attch Xref
                            //string PathName = $"{pathName}\\{dict[name][10]}";
                            string PathName = $"{pathName}\\{sheetObject.xrefName}";
                            //ObjectId acXrefId = db.AttachXref(PathName, dict[name][10]);
                            ObjectId acXrefId = db.AttachXref(PathName, sheetObject.xrefName);
                            
                            if (!acXrefId.IsNull)
                            {
                                // Attach the DWG reference to the current space
                                //Point3d insPt = new Point3d(0, 0, sheetObject.viewCentre.z);
                                Point3d insPt = new Point3d(0, 0, 0);
                                using (BlockReference blockRef = new BlockReference(insPt, acXrefId))
                                {
                                    //blockRef.Layer = "0";
                                    BlockTable blocktable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord modelSpace = trans.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                    modelSpace.AppendEntity(blockRef);

                                    trans.AddNewlyCreatedDBObject(blockRef, true);
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

                                //Point3d revitViewportCentre = new Point3d(double.Parse(dict[name][5]), double.Parse(dict[name][6]), 0);
                                XYZ vpCentre = sheetObject.viewportCentre;
                                Point3d revitViewportCentre = new Point3d(vpCentre.x, vpCentre.y, 0);

                                //Point3d revitViewCentreWCS = new Point3d(double.Parse(dict[name][1]), double.Parse(dict[name][2]), 0);
                                XYZ _revitViewCentreWCS = sheetObject.viewCentre;
                                //Point3d revitViewCentreWCS = new Point3d(revitViewCentre.x, revitViewCentre.y, 0);
                                Point3d revitViewCentreWCS = new Point3d(_revitViewCentreWCS.x, _revitViewCentreWCS.y, 0);

                                //double degrees = DegToRad(double.Parse(dict[name][4]));
                                double degrees = DegToRad(sheetObject.angleToNorth);
                                //double vpWidht = double.Parse(dict[name][8]);
                                double vpWidht = sheetObject.viewportWidth;
                                //double vpHeight = double.Parse(dict[name][9]);
                                double vpHeight = sheetObject.viewportHeight;

                                if (VP != null && CurrentLo.GetViewports().Count == 2) //by default the Layout is a viewport too...https://forums.autodesk.com/t5/net/layouts-and-viewports/td-p/3128748
                                {
                                    UpdateViewport(VP, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                                }
                                else if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < 100)  //Should use the closest viewport, not a fixed distance
                                {
                                    UpdateViewport(VP, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                                }
                            }

                            //Purge unused layers
                            PurgeUnusedLayers(trans, db);

                            trans.Commit();
                        }

                        BindXrefs(db);

                        //currentOpenEditor.Command("_.ZOOM","e");

                        db.Audit(true, true);

                        ed.WriteMessage("\nSaving to file: {0}", outputPath);

                        db.SaveAs(outputPath, DwgVersion.Current);

                        saved++;

                        processed++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("\nProblem processing file: {0} - \"{1}\"", sheetObject.sheetName, ex.Message);

                        problem++;
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


        [CommandMethod("MERGE")]
        public void Merge()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2017/ENU/AutoCAD-NET/files/GUID-FAC1A5EB-2D9E-497B-8FD9-E11D2FF87B93-htm.html

            //https://adndevblog.typepad.com/autocad/2012/07/using-readdwgfile-with-net-attachxref-or-objectarx-acdbattachxref.html
            //Database oldDb = HostApplicationServices.WorkingDatabase; //is it necessary?



            // User should input the folder where the dwgs are saved
            PromptResult pr = ed.GetString("\nEnter folder containing DWGs to process: ");

            if (pr.Status != PromptStatus.OK)
                return;

            string pathName = pr.StringResult;

            string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

            // We'll use some counters to keep track
            // of how the processing is going
            int processed = 0, saved = 0, problem = 0;

            //var dict = File.ReadLines($"{pathName}\\summary.csv").Select(line => line.Split(',')).ToDictionary(line => line[0], line => line.ToList());

            //dict.Remove(dict.Keys.First()); //remove the csv header

            //using a Sheet Object
            var logFile = File.ReadAllLines($"{pathName}\\summary.csv").Select(line => line.Split(',')).ToList<string[]>();
            logFile.RemoveAt(0);

            List<SheetObject> sheetsList = new List<SheetObject>();

            foreach (string[] item in logFile)
            {
                XYZ vc = new XYZ(Convert.ToDouble(item[1]), Convert.ToDouble(item[2]), Convert.ToDouble(item[3]));
                XYZ vpCentre = new XYZ(Convert.ToDouble(item[5]), Convert.ToDouble(item[6]), Convert.ToDouble(item[7]));

                sheetsList.Add(new SheetObject(item[0], vc, Convert.ToDouble(item[4]), vpCentre, Convert.ToDouble(item[8]), Convert.ToDouble(item[9]), item[10]));
            }

            //foreach (string fileName in dict.Keys)
            foreach (SheetObject sheetObject in sheetsList)
            {

                string name = sheetObject.sheetName;
                string filePath = $"{pathName}\\{sheetObject.sheetName}.dwg";
                string outputPath = $"{pathName}\\{sheetObject.sheetName}.dwg";


                //Database db = new Database(false, false);
                Database db = new Database(false, true);
                using (db)
                {
                    try
                    {
                        ed.WriteMessage($"\n\nProcessing file: {filePath} ");

                        db.ReadDwgFile(filePath, FileShare.ReadWrite, true, "");
                        db.CloseInput(true);

                        LayoutManager lm = LayoutManager.Current;

                        lm.CurrentLayout = "Model"; //is it necessary?

                        using (Transaction trans = db.TransactionManager.StartTransaction())
                        {
                            //Attch Xref
                            //string PathName = $"{pathName}\\{dict[name][10]}";
                            string PathName = $"{pathName}\\{sheetObject.xrefName}";
                            //ObjectId acXrefId = db.AttachXref(PathName, dict[name][10]);
                            ObjectId acXrefId = db.AttachXref(PathName, sheetObject.xrefName);

                            if (!acXrefId.IsNull)
                            {
                                // Attach the DWG reference to the current space
                                //Point3d insPt = new Point3d(0, 0, sheetObject.viewCentre.z);
                                Point3d insPt = new Point3d(0, 0, 0);
                                using (BlockReference blockRef = new BlockReference(insPt, acXrefId))
                                {
                                    //blockRef.Layer = "0";
                                    BlockTable blocktable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord modelSpace = trans.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                    modelSpace.AppendEntity(blockRef);

                                    trans.AddNewlyCreatedDBObject(blockRef, true);
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

                                //Point3d revitViewportCentre = new Point3d(double.Parse(dict[name][5]), double.Parse(dict[name][6]), 0);
                                XYZ vpCentre = sheetObject.viewportCentre;
                                Point3d revitViewportCentre = new Point3d(vpCentre.x, vpCentre.y, 0);

                                //Point3d revitViewCentreWCS = new Point3d(double.Parse(dict[name][1]), double.Parse(dict[name][2]), 0);
                                XYZ _revitViewCentreWCS = sheetObject.viewCentre;
                                //Point3d revitViewCentreWCS = new Point3d(revitViewCentre.x, revitViewCentre.y, 0);
                                Point3d revitViewCentreWCS = new Point3d(_revitViewCentreWCS.x, _revitViewCentreWCS.y, 0);

                                //double degrees = DegToRad(double.Parse(dict[name][4]));
                                double degrees = DegToRad(sheetObject.angleToNorth);
                                //double vpWidht = double.Parse(dict[name][8]);
                                double vpWidht = sheetObject.viewportWidth;
                                //double vpHeight = double.Parse(dict[name][9]);
                                double vpHeight = sheetObject.viewportHeight;

                                if (VP != null && CurrentLo.GetViewports().Count == 2) //by default the Layout is a viewport too...https://forums.autodesk.com/t5/net/layouts-and-viewports/td-p/3128748
                                {
                                    UpdateViewport(VP, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                                }
                                else if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < 100)  //Should use the closest viewport, not a fixed distance
                                {
                                    UpdateViewport(VP, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                                }
                            }

                            //Purge unused layers
                            PurgeUnusedLayers(trans, db);

                            trans.Commit();
                        }

                        //BindXrefs(db);

                        //currentOpenEditor.Command("_.ZOOM","e");

                        db.Audit(true, true);

                        ed.WriteMessage("\nSaving to file: {0}", outputPath);

                        db.SaveAs(outputPath, DwgVersion.Current);

                        saved++;

                        processed++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("\nProblem processing file: {0} - \"{1}\"", sheetObject.sheetName, ex.Message);

                        problem++;
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

        //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2016/ENU/AutoCAD-NET/files/GUID-288B4394-C51F-48CC-8B8C-A27873CFFDC1-htm.html
        public void PurgeUnusedLayers(Transaction acTrans, Database acCurDb)
        {
            LayerTable acLyrTbl;

            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;

            // Create an ObjectIdCollection to hold the object ids for each table record
            ObjectIdCollection acObjIdColl = new ObjectIdCollection();

            // Step through each layer and add iterator to the ObjectIdCollection
            foreach (ObjectId acObjId in acLyrTbl)
            {
                acObjIdColl.Add(acObjId);
            }

            // Remove the layers that are in use and return the ones that can be erased
            acCurDb.Purge(acObjIdColl);

            // Step through the returned ObjectIdCollection
            // and erase each unreferenced layer
            foreach (ObjectId acObjId in acObjIdColl)
            {
                SymbolTableRecord acSymTblRec;
                acSymTblRec = acTrans.GetObject(acObjId, OpenMode.ForWrite) as SymbolTableRecord;
                try
                {
                    // Erase the unreferenced layer
                    acSymTblRec.Erase(true);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception Ex)
                {
                    // Layer could not be deleted
                    Application.ShowAlertDialog("Error:\n" + Ex.Message);
                }
            }
        }

        public void UpdateViewport(Viewport _vp, Point3d rvtCentre, Point3d rvtCentreWCS, double degrees, double vpWidth, double vpHeight)
        {
            _vp.UpgradeOpen();
            double cs = _vp.CustomScale; //save the original scale as it changes when we change viewport width and height
            _vp.CenterPoint = rvtCentre; //move the viewport to the revit location
            _vp.Width = vpWidth; //set the width to match the revit width
            _vp.Height = vpHeight; //idem
            _vp.CustomScale = cs;
            //_vp.BackClipOn = true;
            //_vp.FrontClipOn = true;
            //VP.Erase();

            TwistViewport(_vp.Id, rvtCentreWCS, degrees);
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



        // overload no scale factor or set the view - just make the layout
        public void LayoutAndViewport(Database db, string layoutName, out ObjectId rvpid, string deviceName, string mediaName, out ObjectId id)
        {
            // set default values
            rvpid = new ObjectId();
            bool flagVp = false; // flag to create a new floating view port
            double viewSize = (double)Application.GetSystemVariable("VIEWSIZE");
            double height = viewSize;
            double width = viewSize;
            Point2d loCenter = new Point2d(); // layout center point
            Point2d vpLowerCorner = new Point2d();
            Point2d vpUpperCorner = new Point2d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            LayoutManager lm = LayoutManager.Current;
            id = lm.CreateLayout(layoutName);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Layout lo = tr.GetObject(id, OpenMode.ForWrite, false) as Layout;
                if (lo != null)
                {
                    lm.CurrentLayout = lo.LayoutName; // make it current!

                    #region do some plotting settings here for the paper size...
                    ObjectId loid = lm.GetLayoutId(lo.LayoutName);

                    PlotInfo pi = new PlotInfo();
                    pi.Layout = loid;

                    PlotSettings ps = new PlotSettings(false);
                    PlotSettingsValidator psv = PlotSettingsValidator.Current;

                    psv.RefreshLists(ps);
                    psv.SetPlotConfigurationName(ps, deviceName, mediaName);
                    psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
                    psv.SetPlotPaperUnits(ps, PlotPaperUnit.Inches);
                    psv.SetUseStandardScale(ps, true);
                    psv.SetStdScaleType(ps, StdScaleType.ScaleToFit); // use this as default

                    pi.OverrideSettings = ps;

                    PlotInfoValidator piv = new PlotInfoValidator();
                    piv.Validate(pi);

                    lo.CopyFrom(ps);

                    PlotConfig pc = PlotConfigManager.CurrentConfig;
                    // returns data in millimeters...
                    MediaBounds mb = pc.GetMediaBounds(mediaName);

                    Point2d p1 = mb.LowerLeftPrintableArea;
                    Point2d p3 = mb.UpperRightPrintableArea;
                    Point2d p2 = new Point2d(p3.X, p1.Y);
                    Point2d p4 = new Point2d(p1.X, p3.Y);

                    // convert millimeters to inches
                    double mm2inch = 25.4;
                    height = p1.GetDistanceTo(p4) / mm2inch;
                    width = p1.GetDistanceTo(p2) / mm2inch;

                    vpLowerCorner = lo.PlotOrigin;
                    vpUpperCorner = new Point2d(vpLowerCorner.X + width, vpLowerCorner.Y + height);
                    LineSegment2d seg = new LineSegment2d(vpLowerCorner, vpUpperCorner);
                    loCenter = seg.MidPoint;
                    #endregion

                    if (lo.GetViewports().Count == 1) // Viewport was not created by default
                    {
                        // the create by default view ports on new layouts it 
                        // is off we need to mark a flag to generate a new one
                        // in another transaction - out of this one
                        flagVp = true;
                    }
                    else if (lo.GetViewports().Count == 2) // create Viewports by default it is on
                    {
                        // extract the last item from the collection
                        // of view ports inside of the layout
                        int i = lo.GetViewports().Count - 1;
                        ObjectId vpId = lo.GetViewports()[i];

                        if (!vpId.IsNull)
                        {
                            Viewport vp = tr.GetObject(vpId, OpenMode.ForWrite, false) as Viewport;
                            if (vp != null)
                            {
                                vp.Height = height; // change height
                                vp.Width = width; // change width
                                vp.CenterPoint = new Point3d(loCenter.X, loCenter.Y, 0.0); // change center
                                //vp.ColorIndex = 1; // debug

                                // zoom to the Viewport extents
                                Zoom(new Point3d(vpLowerCorner.X, vpLowerCorner.Y, 0.0),
                                    new Point3d(vpUpperCorner.X, vpUpperCorner.Y, 0.0), new Point3d(), 1.0);

                                rvpid = vp.ObjectId; // return the output ObjectId to out...
                            }
                        }
                    }
                }
                tr.Commit();
            } // end of transaction

            // we need another transaction to create a new paper space floating Viewport
            if (flagVp)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr_ps = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

                    Viewport vp = new Viewport();
                    vp.Height = height; // set the height
                    vp.Width = width; // set the width
                    vp.CenterPoint = new Point3d(loCenter.X, loCenter.Y, 0.0); // set the center
                    //vp.ColorIndex = 2; // debug

                    btr_ps.AppendEntity(vp);
                    tr.AddNewlyCreatedDBObject(vp, true);

                    vp.On = true; // make it accessible!

                    // zoom to the Viewport extents
                    Zoom(new Point3d(vpLowerCorner.X, vpLowerCorner.Y, 0.0),
                        new Point3d(vpUpperCorner.X, vpUpperCorner.Y, 0.0), new Point3d(), 1.0);

                    rvpid = vp.ObjectId; // return the ObjectId to the out...

                    tr.Commit();
                } // end of transaction
            }
        }

    }
}