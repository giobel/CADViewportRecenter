using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace TristanAutocadCommands
{
    public class MyCommands
    {
        //Use viewport boundary as selecting window/polygon
        //to find entities in modelspace visible in each viewport
        [CommandMethod("VpSelect")]
        public static void SelectByViewport()
        {
            Document dwg = Application.DocumentManager.MdiActiveDocument;
            Editor ed = dwg.Editor;

            //Save current layout name
            string curLayout = LayoutManager.Current.CurrentLayout;

            try
            {
                //Get viewport information on current layout
                ViewportInfo[] vports = GetViewportInfoOnCurrentLayout();
                if (vports == null) return;

                //Switch to modelspace
                LayoutManager.Current.CurrentLayout = "Model";

                //Select entities in modelspace that are visible
                foreach (ViewportInfo vInfo in vports)
                {
                    ObjectId[] ents = SelectEntitisInModelSpaceByViewport(
                     dwg, vInfo.BoundaryInModelSpace);
                    ed.WriteMessage("\n{0} entit{1} found via Viewport \"{2}\"",
                     ents.Length,
                     ents.Length > 1 ? "ies" : "y",
                     vInfo.ViewportId.ToString());
                }

                Autodesk.AutoCAD.Internal.Utils.PostCommandPrompt();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nCommand \"VpSelect\" failed:");
                ed.WriteMessage("\n{0}\n{1}", ex.Message, ex.StackTrace);
            }
            finally
            {
                //Restore back to original layout
                if (LayoutManager.Current.CurrentLayout != curLayout)
                {
                    LayoutManager.Current.CurrentLayout = curLayout;
                }

                Autodesk.AutoCAD.Internal.Utils.PostCommandPrompt();
            }
        }

        //Determine a given entity in modelspace is visible in
        //which viewports
        [CommandMethod("GetViewports")]
        public static void FindContainingViewport()
        {
            Document dwg = Application.DocumentManager.MdiActiveDocument;
            Editor ed = dwg.Editor;

            //Switch to modelspace
            string curLayout = LayoutManager.Current.CurrentLayout;

            try
            {
                //Get viewport information on current layout
                ViewportInfo[] vports = GetViewportInfoOnCurrentLayout();
                if (vports == null) return;

                //Pick an entity in modelspace
                LayoutManager.Current.CurrentLayout = "Model";
                ObjectId entId = PickEntity(ed);
                if (entId.IsNull)
                {
                    ed.WriteMessage("\n*Cancel*");
                }
                else
                {
                    //Find viewport in which the selected entity is visible
                    List <ObjectId> lst = new List <ObjectId> ();
                    foreach (ViewportInfo vpInfo in vports)
                    {
                        if (IsEntityInsideViewportBoundary(
                          dwg, entId, vpInfo.BoundaryInModelSpace))
                        {
                            lst.Add(vpInfo.ViewportId);
                            ed.WriteMessage(
                            "\nSelected entity is visible in viewport \"{0}\"",
                             vpInfo.ViewportId.ToString());
                        }
                    }

                    if (lst.Count == 0)
                        ed.WriteMessage(
                        "\nSelected entity is not visible in all viewports");
                    else
                        ed.WriteMessage(
                        "\nSelected entity is visible in {0} viewport{1}.",
                         lst.Count, lst.Count > 1 ? "s" : "");
                }

                Autodesk.AutoCAD.Internal.Utils.PostCommandPrompt();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nCommand \"GetViewports\" failed:");
                ed.WriteMessage("\n{0}\n{1}", ex.Message, ex.StackTrace);
            }
            finally
            {
                //Restore back to original layout
                if (LayoutManager.Current.CurrentLayout != curLayout)
                {
                    LayoutManager.Current.CurrentLayout = curLayout;
                }

                Autodesk.AutoCAD.Internal.Utils.PostCommandPrompt();
            }
        }

        public static ViewportInfo[] GetViewportInfoOnCurrentLayout()
        {
            string layoutName = LayoutManager.Current.CurrentLayout;
            if (layoutName.ToUpper() == "MODEL")
            {
                Application.ShowAlertDialog("Please set a layout as active layout!");
                return null;
            }
            else
            {
                Document dwg = Application.DocumentManager.MdiActiveDocument;
                ViewportInfo[] vports =
                 CadHelper.SelectLockedViewportInfoOnLayout(dwg, layoutName);
                if (vports.Length == 0)
                {
                    Application.ShowAlertDialog(
                     "No locked viewport found on layout \"" + layoutName + "\".");
                    return null;
                }
                else
                {
                    return vports;
                }
            }
        }

        public static ObjectId[] SelectEntitisInModelSpaceByViewport(Document dwg, Point3dCollection boundaryInModelSpace)
        {
            ObjectId[] ids = null;

            using (Transaction tran = dwg.TransactionManager.StartTransaction())
            {
                //Zoom to the extents of the viewport boundary in modelspace
                //before calling Editor.SelectXxxxx()
                ZoomToWindow(boundaryInModelSpace);

                PromptSelectionResult res = dwg.Editor.SelectCrossingPolygon(boundaryInModelSpace);
                if (res.Status == PromptStatus.OK)
                {
                    ids = res.Value.GetObjectIds();
                }

                //Restored to previous view (view before zoomming)
                tran.Abort();
            }

            return ids;
        }

        private static void ZoomToWindow(Point3dCollection boundaryInModelSpace)
        {
            Extents3d ext =
              GetViewportBoundaryExtentsInModelSpace(boundaryInModelSpace);

            double[] p1 = new double[] { ext.MinPoint.X, ext.MinPoint.Y, 0.00 };
            double[] p2 = new double[] { ext.MaxPoint.X, ext.MaxPoint.Y, 0.00 };

            dynamic acadApp = Application.AcadApplication;
            acadApp.ZoomWindow(p1, p2);
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

        private static ObjectId PickEntity(Editor ed)
        {
            PromptEntityOptions opt =
             new PromptEntityOptions("\nSelect an entity:");
            PromptEntityResult res = ed.GetEntity(opt);
            if (res.Status == PromptStatus.OK)
            {
                return res.ObjectId;
            }
            else
            {
                return ObjectId.Null;
            }
        }

        private static bool IsEntityInsideViewportBoundary(Document dwg, ObjectId entId, Point3dCollection boundaryInModelSpace)
        {
            bool inside = false;
            using (Transaction tran = dwg.TransactionManager.StartTransaction())
            {
                //Zoom to the extents of the viewport boundary in modelspace
                //before calling Editor.SelectXxxxx()
                ZoomToWindow(boundaryInModelSpace);

                PromptSelectionResult res =
                  dwg.Editor.SelectCrossingPolygon(boundaryInModelSpace);
                if (res.Status == PromptStatus.OK)
                {
                    foreach (ObjectId id in res.Value.GetObjectIds())
                    {
                        if (id == entId)
                        {
                            inside = true;
                            break;
                        }
                    }
                }

                //Restored to previous view (before zoomming)
                tran.Abort();
            }

            return inside;
        }
    }
}
