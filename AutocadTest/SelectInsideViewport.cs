using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Runtime.InteropServices;

namespace AutocadTest
{
    public class SelectInsideViewport
    {
        // https://adndevblog.typepad.com/autocad/2012/04/selecting-model-space-entities-from-paper-space-using-autocad-selection-sets-in-c.html
        // select all entities in Model Space using Paper Space viewport
        // by Fenton Webb, DevTech, Autodesk, 02/Apr/2012
        [CommandMethod("selectMsFromPs", CommandFlags.NoTileMode)]
        static public void selectMsFromPs()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // pick a PS Viewport
            PromptEntityOptions opts = new PromptEntityOptions("Pick PS Viewport");
            opts.SetRejectMessage("Must select PS Viewport objects only");
            opts.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Viewport), false);
            PromptEntityResult res = ed.GetEntity(opts);
            if (res.Status == PromptStatus.OK)
            {
                int vpNumber = 0;
                // extract the viewport points
                Point3dCollection psVpPnts = new Point3dCollection();
                using (Autodesk.AutoCAD.DatabaseServices.Viewport psVp = res.ObjectId.Open(OpenMode.ForRead)
            as Autodesk.AutoCAD.DatabaseServices.Viewport)
                {
                    // get the vp number
                    vpNumber = psVp.Number;
                    // now extract the viewport geometry
                    psVp.GetGripPoints(psVpPnts, new IntegerCollection(), new IntegerCollection());
                }

                // let's assume a rectangular vport for now, make the cross-direction grips square
                Point3d tmp = psVpPnts[2];
                psVpPnts[2] = psVpPnts[1];
                psVpPnts[1] = tmp;

                // Transform the PS points to MS points
                ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 3));
                ResultBuffer rbTo = new ResultBuffer(new TypedValue(5003, 2));
                double[] retPoint = new double[] { 0, 0, 0 };
                // loop the ps points 
                Point3dCollection msVpPnts = new Point3dCollection();
                foreach (Point3d pnt in psVpPnts)
                {
                    // translate from from the DCS of Paper Space (PSDCS) RTSHORT=3 and 
                    // the DCS of the current model space viewport RTSHORT=2
                    acedTrans(pnt.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, retPoint);
                    // add the resulting point to the ms pnt array
                    msVpPnts.Add(new Point3d(retPoint));
                    ed.WriteMessage("\n" + new Point3d(retPoint).ToString());
                }

                // now switch to MS
                ed.SwitchToModelSpace();
                // set the CVPort
                Application.SetSystemVariable("CVPORT", vpNumber);
                // once switched, we can use the normal selection mode to select
                PromptSelectionResult selectionresult = ed.SelectCrossingPolygon(msVpPnts);
                // now switch back to PS
                ed.SwitchToPaperSpace();
            }
        }

        //https://adndevblog.typepad.com/autocad/2012/03/converting-paper-space-point-to-model-space-using-autocadnet.html
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);

    }
}
