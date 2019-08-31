﻿using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutocadTest
{
    public class CommandsMergeCopy
    {
        //http://drive-cad-with-code.blogspot.com/2014/03/selecting-entities-in-modelspace.html

        [CommandMethod("MERGEMERGE")]
        public void MergeMerge()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Layout1";

            Database db = doc.Database;

            string dwgName = Path.GetFileNameWithoutExtension(doc.Name);

            string folderPath = Path.GetDirectoryName(doc.Name);

            //using a Sheet Object
            var logFile = File.ReadAllLines($"{folderPath}\\summary.csv").Select(line => line.Split(',')).ToList<string[]>();
            logFile.RemoveAt(0);

            //get csv file content
            List<SheetObject> sheetsList = new List<SheetObject>();
            foreach (string[] item in logFile)
            {
                XYZ vc = new XYZ(Convert.ToDouble(item[1]), Convert.ToDouble(item[2]), Convert.ToDouble(item[3]));
                XYZ vpCentre = new XYZ(Convert.ToDouble(item[5]), Convert.ToDouble(item[6]), Convert.ToDouble(item[7]));

                sheetsList.Add(new SheetObject(item[0], vc, Convert.ToDouble(item[4]), vpCentre, Convert.ToDouble(item[8]), Convert.ToDouble(item[9]), item[10]));
            }

            
            SheetObject sheetObject = sheetsList.Where(x => x.sheetName == dwgName).First();

            XYZ currentVpCentre = sheetObject.viewportCentre;

            Point3d revitViewportCentre = new Point3d(currentVpCentre.x, currentVpCentre.y, 0);

            XYZ _revitViewCentreWCS = sheetObject.viewCentre;

            Point3d revitViewCentreWCS = new Point3d(_revitViewCentreWCS.x, _revitViewCentreWCS.y, 0);

            double degrees = Helpers.DegToRad(sheetObject.angleToNorth);

            double vpWidht = sheetObject.viewportWidth;

            double vpHeight = sheetObject.viewportHeight;


            //get document name
            ed.WriteMessage("\n======================== Dwg Name: " + doc.Name + "\n");
            ed.WriteMessage("======================== Xref(s): " + sheetObject.xrefName + "\n");

            //find the xref viewport and delete its content
            //ed.SwitchToPaperSpace();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                string currentLo = lm.CurrentLayout;

                DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;

                Viewport matchingViewport = null;

                List<ObjectId> layerToFreeze = new List<ObjectId>();


                //Create Layer to store xref
                #region
                string layerName = $"0-{sheetObject.xrefName}";

                Helpers.CreateLayer(db, trans, layerName);

                ed.WriteMessage("======================== Layer created: " + layerName + "\n");

                LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                ObjectId layer = new ObjectId();


                foreach (ObjectId layerId in layerTable)
                {
                    LayerTableRecord currentLayer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                    if (currentLayer.Name == layerName)
                    {
                        layer = layerId;
                        layerToFreeze.Add(layerId);
                    }
                }

                #endregion

                //Find the equivalent Revit viewport
                #region
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

                //Delete Viewport Content
                #region
                Point3dCollection vpCorners = GetViewportBoundary(matchingViewport);

                Matrix3d mt = PaperToModel(matchingViewport);

                Point3dCollection vpCornersInModel = TransformPaperSpacePointToModelSpace(vpCorners, mt);

                ObjectId[] viewportContent = null;

                try
                {
                    viewportContent = SelectEntitisInModelSpaceByViewport(doc, vpCornersInModel);
                    ed.WriteMessage("======================== Viewport objects: " + viewportContent.Length.ToString() + "\n");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("======================== Error: " + ex.Message + "\n");
                }

                if (viewportContent != null)
                {
                    foreach (ObjectId item in viewportContent)
                    {
                        Entity e = (Entity)trans.GetObject(item, OpenMode.ForWrite);
                        //ed.WriteMessage(item.GetType().Name);
                        e.Erase();
                    }
                    ed.WriteMessage("======================== Viewport content deleted\n");
                }
                else
                {
                    ed.WriteMessage("======================== viewport content is null!\n");
                }

                #endregion

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
                        BlockTable blocktable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = trans.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        modelSpace.AppendEntity(blockRef);

                        trans.AddNewlyCreatedDBObject(blockRef, true);
                    }

                    ed.WriteMessage("======================== xref loaded\n");
                }

                #endregion

                //Recenter Viewport
                #region
                Helpers.UpdateViewport(matchingViewport, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                ed.WriteMessage("======================== Viewport updated \n");
                #endregion

                ed.WriteMessage("Swith to Model layout \n");
                lm.CurrentLayout = "Model";

                //ed.WriteMessage("======================== Run Set by layer\n");
                //ed.Command("-setbylayer", "all"," ","y","y");

                ed.WriteMessage("======================== Run Audit\n");
                ed.Command("audit", "y");

                ed.WriteMessage("======================== Run Purge \n");
                ed.Command("-purge", "all", " ", "n");

                trans.Commit();
            }

            ed.WriteMessage("Switch back to paper space \n");
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