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
    public class CommandRecenterViewports
    {
        [CommandMethod("RECENTERVIEWPORTS")]
        public void RecenterViewports()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== +++ Command Recenter Viewports +++");

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Layout1";

            Database db = doc.Database;

            string dwgName = Path.GetFileNameWithoutExtension(doc.Name);

            string folderPath = Path.GetDirectoryName(doc.Name);

            List<SheetObject> sheetObjects = Helpers.SheetsObjectsFromCSV(folderPath, dwgName);

            //get document name
            ed.WriteMessage("\n=== Dwg Name: " + doc.Name + "\n");


            foreach (SheetObject sheetObject in sheetObjects)
            {
                ed.WriteMessage("=== Xref(s): " + sheetObject.xrefName + "\n");

                XYZ currentVpCentre = sheetObject.viewportCentre;

                Point3d revitViewportCentre = new Point3d(currentVpCentre.x, currentVpCentre.y, 0);

                XYZ _revitViewCentreWCS = sheetObject.viewCentre;

                Point3d revitViewCentreWCS = new Point3d(_revitViewCentreWCS.x, _revitViewCentreWCS.y, 0);

                double degrees = Helpers.DegToRad(sheetObject.angleToNorth);

                double vpWidht = sheetObject.viewportWidth;

                double vpHeight = sheetObject.viewportHeight;

                string layerName = $"0-{sheetObject.xrefName}";

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    string currentLo = lm.CurrentLayout;

                    DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                    Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;

                    Viewport matchingViewport = null;

                    LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                    List<ObjectId> layerToFreeze = new List<ObjectId>();

                    foreach (ObjectId layerId in layerTable)
                    {
                        LayerTableRecord currentLayer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        if (currentLayer.Name == layerName)
                        {
                            layerToFreeze.Add(layerId);
                        }
                    }

                    //Find the equivalent Revit viewport
                    #region
                    foreach (ObjectId ID in CurrentLo.GetViewports())
                    {
                        Viewport VP = trans.GetObject(ID, OpenMode.ForWrite) as Viewport;

                        if (VP != null && CurrentLo.GetViewports().Count == 2 && VP.CenterPoint.X > 20) //by default the Layout is a viewport too...https://forums.autodesk.com/t5/net/layouts-and-viewports/td-p/3128748
                        {
                            matchingViewport = VP;
                            ed.WriteMessage("=== Single Viewport on sheet\n");
                        }
                        if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < Helpers.ViewportDistanceTolerance)  //Should use the closest viewport, not a fixed distance
                        {
                            matchingViewport = VP;
                            ed.WriteMessage("=== Multiple Viewports on sheet\n");
                        }
                        else
                        {
                            VP.FreezeLayersInViewport(layerToFreeze.GetEnumerator());
                        }
                    }
                    ed.WriteMessage("=== Viewport Name: " + matchingViewport.BlockName + "\n");
                    ed.WriteMessage("=== Viewport Center: " + matchingViewport.CenterPoint + "\n");
                    #endregion

                    Helpers.UpdateViewport(matchingViewport, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                    ed.WriteMessage("=== Viewport updated \n");

                    trans.Commit();
                }//close transaction

            }

            ed.WriteMessage("Save file \n");
            db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

            ed.WriteMessage("\n=== +++ Command Recenter Viewports End +++");
        }


    }//close class
}//close namespace