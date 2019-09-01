using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TristanAutocadCommands
{
    class Helpers
    {

        //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2016/ENU/AutoCAD-NET/files/GUID-288B4394-C51F-48CC-8B8C-A27873CFFDC1-htm.html
        public static void PurgeUnusedLayers(Transaction acTrans, Database acCurDb)
        {
            LayerTable acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;


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


        //https://www.keanw.com/2007/08/purging-registe.html
        public static int PurgeDatabase(Database db, Transaction tr)
        {
            int idCount = 0;
            
                // Create the list of objects to "purge"
                ObjectIdCollection idsToPurge = new ObjectIdCollection();
                
                // Add all the Registered Application names
                
                RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId,OpenMode.ForRead);


                foreach (ObjectId raId in rat)

                {

                    if (raId.IsValid)

                    {

                        idsToPurge.Add(raId);

                    }

                }


                // Call the Purge function to filter the list


                db.Purge(idsToPurge);


                // Erase each of the objects we've been

                // allowed to


                foreach (ObjectId id in idsToPurge)

                {

                    DBObject obj =

                      tr.GetObject(id, OpenMode.ForWrite);


                    // Let's just add to me "debug" code

                    // to list the registered applications

                    // we're erasing


                    RegAppTableRecord ratr =

                      obj as RegAppTableRecord;


                    obj.Erase();

                }


                // Return the number of objects erased

                // (i.e. purged)


                idCount = idsToPurge.Count;           

            return idCount;
    }

    public static void UpdateViewport(Viewport _vp, Point3d rvtCentre, Point3d rvtCentreWCS, double degrees, double vpWidth, double vpHeight)
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

        public static void BindXrefs(Database db)

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

        public static double DegToRad(double deg)
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

        public static void CreateLayer(Database acCurDb, Transaction acTrans, string sLayerName)
        {
            // Open the Layer table for read
            LayerTable layerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
            if (layerTable.Has(sLayerName) == false)
            {
                using (LayerTableRecord acLyrTblRec = new LayerTableRecord())
                {
                    // Assign the layer the ACI color 3 and a name
                    acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                    acLyrTblRec.Name = sLayerName;

                    // Upgrade the Layer table for write
                    layerTable.UpgradeOpen();

                    // Append the new layer to the Layer table and the transaction
                    layerTable.Add(acLyrTblRec);
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                }
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
