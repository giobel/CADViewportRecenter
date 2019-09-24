using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MxAutocadBatchCommands
{
    public class CommandsMergeMultiple
    {
        //http://drive-cad-with-code.blogspot.com/2014/03/selecting-entities-in-modelspace.html

        [CommandMethod("MERGEMULTIPLE")]
        public void Merge()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Layout1";

            Database db = doc.Database;

            string dwgName = Path.GetFileNameWithoutExtension(doc.Name);

            string folderPath = Path.GetDirectoryName(doc.Name);

            List<SheetObject> sheetObjects = Helpers.SheetsObjectsFromCSV(folderPath, dwgName);



            using (Transaction transCreateLayers = db.TransactionManager.StartTransaction())
            {
                foreach (SheetObject sheetObject in sheetObjects)
                {
                    string layerName = $"0-{sheetObject.xrefName}";

                    Helpers.CreateLayer(db, transCreateLayers, layerName);

                    ed.WriteMessage("======================== Layer created: " + layerName + "\n");

                    LayerTable layerTable = transCreateLayers.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                    ObjectId layer = new ObjectId();


                    foreach (ObjectId layerId in layerTable)
                    {
                        LayerTableRecord currentLayer = transCreateLayers.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        if (currentLayer.Name == layerName)
                        {
                            layer = layerId;
                        }
                    }


                    //Load Xref
                    #region
                    string PathName = $"{folderPath}\\{sheetObject.xrefName}";

                    ObjectId acXrefId = db.AttachXref(PathName, sheetObject.xrefName);

                    if (!acXrefId.IsNull)
                    {
                        // Attach the DWG reference to the model space
                        Point3d insPt = new Point3d(0, 0, 0);
                        using (BlockReference blockRef = new BlockReference(insPt, acXrefId))
                        {
                            blockRef.SetLayerId(layer, true);
                            BlockTable blocktable = transCreateLayers.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                            BlockTableRecord modelSpace = transCreateLayers.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                            modelSpace.AppendEntity(blockRef);

                            transCreateLayers.AddNewlyCreatedDBObject(blockRef, true);
                        }

                        ed.WriteMessage("======================== xref loaded\n");
                    }

                    #endregion
                }

                transCreateLayers.Commit();
            }
            //get document name
            ed.WriteMessage("\n======================== Dwg Name: " + doc.Name + "\n");

            //objects to delete
            List<ObjectId[]> viewportContentList = new List<ObjectId[]>();

            foreach (SheetObject sheetObject in sheetObjects)
                {

                string layerName = $"0-{sheetObject.xrefName}";

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    ed.WriteMessage("======================== Xref(s): " + sheetObject.xrefName + "\n");

                    XYZ currentVpCentre = sheetObject.viewportCentre;

                    Point3d revitViewportCentre = new Point3d(currentVpCentre.x, currentVpCentre.y, 0);

                    XYZ _revitViewCentreWCS = sheetObject.viewCentre;

                    Point3d revitViewCentreWCS = new Point3d(_revitViewCentreWCS.x, _revitViewCentreWCS.y, 0);

                    double degrees = Helpers.DegToRad(sheetObject.angleToNorth);

                    double vpWidht = sheetObject.viewportWidth;

                    double vpHeight = sheetObject.viewportHeight;

                    lm.CurrentLayout = "Layout1";

                    string currentLo = lm.CurrentLayout;

                    DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                    Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;

                    Viewport matchingViewport = null;

                    List<ObjectId> layerToFreeze = new List<ObjectId>();


                    //LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                    //foreach (ObjectId layerId in layerTable)
                    //{
                    //    LayerTableRecord currentLayer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                    //    if (currentLayer.Name != layerName)
                    //    {
                    //        layerToFreeze.Add(layerId);
                    //    }
                    //}

                    #region Find the equivalent Revit viewport
                    foreach (ObjectId ID in CurrentLo.GetViewports())
                    {
                        Viewport VP = trans.GetObject(ID, OpenMode.ForWrite) as Viewport;

                        if (VP != null && CurrentLo.GetViewports().Count == 2 && VP.CenterPoint.X > 20) //by default the Layout is a viewport too...https://forums.autodesk.com/t5/net/layouts-and-viewports/td-p/3128748
                        {
                            matchingViewport = VP;
                            ed.WriteMessage("======================== Single Viewport on sheet\n");
                        }
                        if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < 100)  //Should use the closest viewport, not a fixed distance
                        {
                            matchingViewport = VP;
                            ed.WriteMessage("======================== Multiple Viewports on sheet\n");
                        }
                        else
                        {
                            VP.FreezeLayersInViewport(layerToFreeze.GetEnumerator());
                        }
                    }
                    ed.WriteMessage("======================== Viewport Name: " + matchingViewport.BlockName + "\n");
                    ed.WriteMessage("======================== Viewport Center: " + matchingViewport.CenterPoint + "\n");
                    #endregion


                    #region Delete Viewport Content
                    Point3dCollection vpCorners = GetViewportBoundary(matchingViewport);

                    Matrix3d mt = PaperToModel(matchingViewport);

                    Point3dCollection vpCornersInModel = TransformPaperSpacePointToModelSpace(vpCorners, mt);

                    try
                    {
                        ObjectId[] viewportContent = SelectEntitisInModelSpaceByViewport(doc, vpCornersInModel);
                        viewportContentList.Add(viewportContent);
                        ed.WriteMessage("======================== Viewport objects: " + viewportContent.Length.ToString() + "\n");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("======================== Error: " + ex.Message + "\n");
                    }
                    #endregion

                    //Recenter Viewport
                    #region
                    Helpers.UpdateViewport(matchingViewport, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                    ed.WriteMessage("======================== Viewport updated \n");
                    #endregion

                    trans.Commit();

                }//close transaction
            }//close loop through sheet objects


            using (Transaction transDeleteObjects = db.TransactionManager.StartTransaction())
            {


                foreach (ObjectId[] itemList in viewportContentList)
                {
                    if (itemList != null)
                    {
                        foreach (ObjectId item in itemList)
                        {
                            Entity e = (Entity)transDeleteObjects.GetObject(item, OpenMode.ForWrite);
                            //ed.WriteMessage(item.GetType().Name);
                            e.Erase();
                        }
                        ed.WriteMessage("======================== Viewport content deleted\n");
                    }
                    else
                    {
                        ed.WriteMessage("======================== viewport content is null!\n");
                    }
                }
                transDeleteObjects.Commit();
            }


            ed.WriteMessage("Switch to Model layout \n");
            lm.CurrentLayout = "Model";

            //ed.WriteMessage("======================== Run Set by layer\n");
            //ed.Command("-setbylayer", "all"," ","y","y");

            ed.WriteMessage("======================== Run Audit\n");
            ed.Command("audit", "y");

            ed.WriteMessage("======================== Run Purge \n");
            ed.Command("-purge", "all", " ", "n");

            lm.CurrentLayout = "Layout1";

            ed.Command("_.zoom", "_extents");
            ed.Command("_.zoom", "_scale", "0.9");

            ed.WriteMessage("Save file \n");
            db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

            ed.WriteMessage("done");
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

        private static Point3dCollection TransformPaperSpacePointToModelSpace(Point3dCollection paperSpacePoints, Matrix3d mt)
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

    }//close class
}//close namespace