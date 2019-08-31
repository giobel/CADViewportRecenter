using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AutocadTest
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
                    //Application.ShowAlertDialog("Error:\n" + Ex.Message);
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

            RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);


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

        public static List<SheetObject> SheetsObjectsFromCSV(string folderPath, string SheetNumber)
        {
            var logFile = File.ReadAllLines($"{folderPath}\\summary.csv").Select(line => line.Split(',')).ToList<string[]>();
            logFile.RemoveAt(0);

            List<SheetObject> sheetsList = new List<SheetObject>();
            foreach (string[] item in logFile)
            {
                XYZ vc = new XYZ(Convert.ToDouble(item[1]), Convert.ToDouble(item[2]), Convert.ToDouble(item[3]));
                XYZ vpCentre = new XYZ(Convert.ToDouble(item[5]), Convert.ToDouble(item[6]), Convert.ToDouble(item[7]));

                sheetsList.Add(new SheetObject(item[0], vc, Convert.ToDouble(item[4]), vpCentre, Convert.ToDouble(item[8]), Convert.ToDouble(item[9]), item[10], item[11]));
            }

            return sheetsList.Where(x => x.sheetName == SheetNumber).ToList();

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

        public static ObjectId[] SelectEntitisInModelSpaceByViewport(Document doc, Point3dCollection boundaryInModelSpace)
        {
            //doc.Editor.SwitchToModelSpace();
            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Model";

            ObjectId[] ids = null;

            PromptSelectionResult res = doc.Editor.SelectCrossingPolygon(boundaryInModelSpace);
            if (res.Status == PromptStatus.OK)
            {
                ids = res.Value.GetObjectIds();
            }


            //doc.Editor.SwitchToPaperSpace();
            return ids;
        }

        public static Point3dCollection GetViewportBoundary(Viewport vport)
        {
            Point3dCollection points = new Point3dCollection();

            Extents3d ext = vport.GeometricExtents;
            points.Add(new Point3d(ext.MinPoint.X, ext.MinPoint.Y, 0.0));
            points.Add(new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0.0));
            points.Add(new Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, 0.0));
            points.Add(new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0.0));

            return points;
        }

        private static Extents3d GetViewportBoundaryExtentsInModelSpace(Point3dCollection points)
        {
            Extents3d ext = new Extents3d();
            foreach (Point3d p in points)
            {
                ext.AddPoint(p);
            }

            return ext;
        }

        public static Point3dCollection TransformPaperSpacePointToModelSpace(Point3dCollection paperSpacePoints, Matrix3d mt)
        {
            Point3dCollection points = new Point3dCollection();

            foreach (Point3d p in paperSpacePoints)
            {
                points.Add(p.TransformBy(mt));
            }

            return points;
        }

        #region
        //**********************************************************************
        //Create coordinate transform matrix
        //between modelspace and paperspace viewport
        //The code is borrowed from
        //http://www.theswamp.org/index.php?topic=34590.msg398539#msg398539
        //*********************************************************************
        public static Matrix3d PaperToModel(Viewport vp)
        {
            Matrix3d mx = ModelToPaper(vp);
            return mx.Inverse();
        }

        public static Matrix3d ModelToPaper(Viewport vp)
        {
            Vector3d vd = vp.ViewDirection;
            Point3d vc = new Point3d(vp.ViewCenter.X, vp.ViewCenter.Y, 0);
            Point3d vt = vp.ViewTarget;
            Point3d cp = vp.CenterPoint;
            double ta = -vp.TwistAngle;
            double vh = vp.ViewHeight;
            double height = vp.Height;
            double width = vp.Width;
            double scale = vh / height;
            double lensLength = vp.LensLength;
            Vector3d zaxis = vd.GetNormal();
            Vector3d xaxis = Vector3d.ZAxis.CrossProduct(vd);
            Vector3d yaxis;

            if (!xaxis.IsZeroLength())
            {
                xaxis = xaxis.GetNormal();
                yaxis = zaxis.CrossProduct(xaxis);
            }
            else if (zaxis.Z < 0)
            {
                xaxis = Vector3d.XAxis * -1;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis * -1;
            }

            else
            {
                xaxis = Vector3d.XAxis;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis;
            }
            Matrix3d pcsToDCS = Matrix3d.Displacement(Point3d.Origin - cp);
            pcsToDCS = pcsToDCS * Matrix3d.Scaling(scale, cp);
            Matrix3d dcsToWcs = Matrix3d.Displacement(vc - Point3d.Origin);
            Matrix3d mxCoords = Matrix3d.AlignCoordinateSystem(
             Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis,
             Vector3d.ZAxis, Point3d.Origin,
             xaxis, yaxis, zaxis);
            dcsToWcs = mxCoords * dcsToWcs;
            dcsToWcs = Matrix3d.Displacement(vt - Point3d.Origin) * dcsToWcs;
            dcsToWcs = Matrix3d.Rotation(ta, zaxis, vt) * dcsToWcs;

            Matrix3d perspectiveMx = Matrix3d.Identity;
            if (vp.PerspectiveOn)
            {
                double vSize = vh;
                double aspectRatio = width / height;
                double adjustFactor = 1.0 / 42.0;
                double adjstLenLgth = vSize * lensLength *
                 System.Math.Sqrt(1.0 + aspectRatio * aspectRatio) * adjustFactor;
                double iDist = vd.Length;
                double lensDist = iDist - adjstLenLgth;
                double[] dataAry = new double[]
                 {
         1,0,0,0,0,1,0,0,0,0,
         (adjstLenLgth-lensDist)/adjstLenLgth,
         lensDist*(iDist-adjstLenLgth)/adjstLenLgth,
         0,0,-1.0/adjstLenLgth,iDist/adjstLenLgth
                 };

                perspectiveMx = new Matrix3d(dataAry);
            }

            Matrix3d finalMx =
             pcsToDCS.Inverse() * perspectiveMx * dcsToWcs.Inverse();

            return finalMx;
        }

        #endregion

    }
}
