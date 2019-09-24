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
    public class CommandDeleteVP
    {
        //http://drive-cad-with-code.blogspot.com/2014/03/selecting-entities-in-modelspace.html

        [CommandMethod("DELETEVP")]
        public void DeleteVP()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== +++ Command Delete Viewport Content Start +++");

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Layout1";

            Database db = doc.Database;

            string dwgName = Path.GetFileNameWithoutExtension(doc.Name);

            string folderPath = Path.GetDirectoryName(doc.Name);

            List<SheetObject> sheetObjects = Helpers.SheetsObjectsFromCSV(folderPath, dwgName);

            ed.WriteMessage("\n=== Sheets: " + sheetObjects.Count + "\n");

            List<ObjectId[]> viewportContentList = new List<ObjectId[]>();
            //get document name
            ed.WriteMessage("\n=== Dwg Name: " + doc.Name + "\n");


            foreach (SheetObject sheetObject in sheetObjects)
            {
                ed.WriteMessage("=== Xref(s): " + sheetObject.xrefName + "\n");

                XYZ currentVpCentre = sheetObject.viewportCentre;

                Point3d revitViewportCentre = new Point3d(currentVpCentre.x, currentVpCentre.y, 0);


                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    string currentLo = lm.CurrentLayout;

                    DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                    Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;

                    Viewport matchingViewport = null;

                    //Find the equivalent Revit viewport
                    #region
                    foreach (ObjectId ID in CurrentLo.GetViewports())
                    {
                        Viewport VP = trans.GetObject(ID, OpenMode.ForWrite) as Viewport;

                        if (VP != null && CurrentLo.GetViewports().Count == 2 && VP.CenterPoint.X > 20) //by default the Layout is a viewport too...https://forums.autodesk.com/t5/net/layouts-and-viewports/td-p/3128748
                        {
                            matchingViewport = VP;
                            ed.WriteMessage($"=== Single Viewport on sheet {sheetObject.sheetName}\n");
                        }
                        if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < Helpers.ViewportDistanceTolerance)  //Should use the closest viewport, not a fixed distance
                        {
                            matchingViewport = VP;
                            ed.WriteMessage($"=== Multiple Viewports on sheet {sheetObject.sheetName}\n");
                        }
                        else
                        {
                            //VP.FreezeLayersInViewport(layerToFreeze.GetEnumerator());
                        }
                    }
                    ed.WriteMessage("=== Viewport Name: " + matchingViewport.BlockName + "\n");
                    ed.WriteMessage("=== Viewport Center: " + matchingViewport.CenterPoint + "\n");
                    #endregion

                    //Delete Viewport Content
                    #region
                    Point3dCollection vpCorners = Helpers.GetViewportBoundary(matchingViewport);

                    Matrix3d mt = Helpers.PaperToModel(matchingViewport);

                    Point3dCollection vpCornersInModel = Helpers.TransformPaperSpacePointToModelSpace(vpCorners, mt);

                    try
                    {
                        ObjectId[] viewportContent = Helpers.SelectEntitisInModelSpaceByViewport(doc, vpCornersInModel);
                        viewportContentList.Add(viewportContent);
                        ed.WriteMessage("=== Viewport objects to be deleted: " + viewportContent.Length.ToString() + "\n");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("=== Error: " + ex.Message + "\n");
                    }

                    #endregion

                    //trans.Commit();
                }//close transaction

            }


            using (Transaction transDelete = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId[] itemList in viewportContentList)
                {
                    if (itemList != null)
                    {
                        foreach (ObjectId item in itemList)
                        {
                            Entity e = (Entity)transDelete.GetObject(item, OpenMode.ForWrite);
                            //ed.WriteMessage(item.GetType().Name);
                            e.Erase();
                        }
                        ed.WriteMessage("=== Viewport content deleted\n");
                    }
                    else
                    {
                        ed.WriteMessage("=== Viewport content is null!\n");
                    }
                }

                transDelete.Commit();
            }


            ed.Command("_.zoom", "_extents");
            ed.Command("_.zoom", "_scale", "0.9");

            ed.WriteMessage("Save file \n");
            db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

            ed.WriteMessage("=== +++ Command Delete Viewport Content End +++");
        }

        

    }//close class
}//close namespace