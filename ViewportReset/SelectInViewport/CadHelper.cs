﻿using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;

namespace TristanAutocadCommands
{
    //Class to hold Viewport information, obtained
    //in single Transaction
    public class ViewportInfo
    {
        public ObjectId ViewportId { set; get; }
        public ObjectId NonRectClipId { set; get; }
        public Point3dCollection BoundaryInPaperSpace { set; get; }
        public Point3dCollection BoundaryInModelSpace { set; get; }
    }

    public class CadHelper
    {
        //Get needed Viewport information
        public static ViewportInfo[] SelectLockedViewportInfoOnLayout(Document dwg, string layoutName)
        {
            List<ViewportInfo> lst = new List<ViewportInfo>();
            TypedValue[] vals = new TypedValue[] { new TypedValue((int)DxfCode.Start, "VIEWPORT"), new TypedValue((int)DxfCode.LayoutName, layoutName) };

            PromptSelectionResult res =
              dwg.Editor.SelectAll(new SelectionFilter(vals));
            if (res.Status == PromptStatus.OK)
            {
                using (Transaction tran =
                  dwg.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in res.Value.GetObjectIds())
                    {
                        Viewport vport = (Viewport)tran.GetObject(
                          id, OpenMode.ForRead);
                        if (vport.Number != 1 && vport.Locked)
                        {
                            ViewportInfo vpInfo = new ViewportInfo();
                            vpInfo.ViewportId = id;
                            vpInfo.NonRectClipId = vport.NonRectClipEntityId;
                            if (!vport.NonRectClipEntityId.IsNull &&
                              vport.NonRectClipOn)
                            {
                                Polyline2d pl = (Polyline2d)tran.GetObject(
                                  vport.NonRectClipEntityId, OpenMode.ForRead);
                                vpInfo.BoundaryInPaperSpace =
                                 GetNonRectClipBoundary(pl, tran);
                            }
                            else
                            {
                                vpInfo.BoundaryInPaperSpace =
                                 GetViewportBoundary(vport);
                            }

                            Matrix3d mt = PaperToModel(vport);
                            vpInfo.BoundaryInModelSpace =
                             TransformPaperSpacePointToModelSpace(
                             vpInfo.BoundaryInPaperSpace, mt);

                            lst.Add(vpInfo);
                        }
                    }

                    tran.Commit();
                }
            }

            return lst.ToArray();
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

        private static Point3dCollection GetNonRectClipBoundary(
         Polyline2d polyline, Transaction tran)
        {
            Point3dCollection points = new Point3dCollection();

            foreach (ObjectId vxId in polyline)
            {
                Vertex2d vx = (Vertex2d)tran.GetObject(vxId, OpenMode.ForRead);
                points.Add(polyline.VertexPosition(vx));
            }

            return points;
        }

        public static Point3dCollection TransformPaperSpacePointToModelSpace(
         Point3dCollection paperSpacePoints, Matrix3d mt)
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
                 Math.Sqrt(1.0 + aspectRatio * aspectRatio) * adjustFactor;
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

        public static ObjectId[] SelectEntitisInModelSpaceByViewport(Document doc, Point3dCollection boundaryInModelSpace)
        {
            doc.Editor.SwitchToModelSpace();

            ObjectId[] ids = null;
            PromptSelectionResult res = null;

            try
            {
                res = doc.Editor.SelectCrossingPolygon(boundaryInModelSpace);
            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

            if (res.Status == PromptStatus.OK)
            {
                ids = res.Value.GetObjectIds();
            }


            doc.Editor.SwitchToPaperSpace();
            return ids;
        }

    }
}
