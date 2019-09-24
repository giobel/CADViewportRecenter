using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using System;
using System.Linq;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;

namespace MxAutocadCommands
{
    public class CommandMergeAndBind
    {
        [CommandMethod("MERGEANDBIND")]
        public void MergeAndBind()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2017/ENU/AutoCAD-NET/files/GUID-FAC1A5EB-2D9E-497B-8FD9-E11D2FF87B93-htm.html

            //https://adndevblog.typepad.com/autocad/2012/07/using-readdwgfile-with-net-attachxref-or-objectarx-acdbattachxref.html
            
            
            // User should input the folder where the dwgs are saved
            PromptResult pr = ed.GetString("\nEnter folder containing DWGs to process: ");

            if (pr.Status != PromptStatus.OK)
                return;

            string pathName = pr.StringResult;

            string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

            // We'll use some counters to keep track of how the processing is going
            int processed = 0, saved = 0, problem = 0;

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

                        string layerName = $"0-{sheetObject.xrefName}";

                        using (Transaction trans = db.TransactionManager.StartTransaction())
                        {
                            Helpers.CreateLayer(db, trans, layerName);

                            LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                            ObjectId layer = new ObjectId();
                            List<ObjectId> layerToFreeze = new List<ObjectId>();

                            foreach (ObjectId layerId in layerTable)
                            {
                                LayerTableRecord currentLayer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                                if (currentLayer.Name == layerName)
                                {
                                    layer = layerId;
                                    layerToFreeze.Add(layerId);
                                }
                                //else
                                //{
                                //    layerToFreeze.Add(layerId);
                                //}
                            }

                            //Attch Xref
                            string PathName = $"{pathName}\\{sheetObject.xrefName}";

                            ObjectId acXrefId = db.AttachXref(PathName, sheetObject.xrefName);
                            
                            if (!acXrefId.IsNull)
                            {
                                // Attach the DWG reference to the current space
                                Point3d insPt = new Point3d(0, 0, 0);
                                using (BlockReference blockRef = new BlockReference(insPt, acXrefId))
                                {
                                    blockRef.SetLayerId(layer, true);
                                    BlockTable blocktable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord modelSpace = trans.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                    modelSpace.AppendEntity(blockRef);

                                    trans.AddNewlyCreatedDBObject(blockRef, true);
                                }
                            }

                            lm.CurrentLayout = "Layout1";

                            string currentLo = lm.CurrentLayout;

                            DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                            Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;
                            
                            foreach (ObjectId ID in CurrentLo.GetViewports())
                            {
                                Viewport VP = trans.GetObject(ID, OpenMode.ForWrite) as Viewport;

                                XYZ vpCentre = sheetObject.viewportCentre;
                                Point3d revitViewportCentre = new Point3d(vpCentre.x, vpCentre.y, 0);

                                XYZ _revitViewCentreWCS = sheetObject.viewCentre;
                                Point3d revitViewCentreWCS = new Point3d(_revitViewCentreWCS.x, _revitViewCentreWCS.y, 0);

                                double degrees = Helpers.DegToRad(sheetObject.angleToNorth);
                                double vpWidht = sheetObject.viewportWidth;
                                double vpHeight = sheetObject.viewportHeight;

                                if (VP != null && CurrentLo.GetViewports().Count == 2) //by default the Layout is a viewport too...https://forums.autodesk.com/t5/net/layouts-and-viewports/td-p/3128748
                                {
                                    Helpers.UpdateViewport(VP, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                                }
                                if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < 100)  //Should use the closest viewport, not a fixed distance
                                {
                                    Helpers.UpdateViewport(VP, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                                }
                                else
                                {
                                    VP.FreezeLayersInViewport(layerToFreeze.GetEnumerator());
                                }
                            }
                            //Purge unused layers
                            Helpers.PurgeUnusedLayers(trans, db);

                            Helpers.PurgeDatabase(db, trans);

                            trans.Commit();                        
                        }
                        Helpers.BindXrefs(db);

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
    }
}