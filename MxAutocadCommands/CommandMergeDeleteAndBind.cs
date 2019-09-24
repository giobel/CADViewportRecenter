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
    public class CommandMergeDeleteAndBind
    {
        [CommandMethod("MERGEDELETEANDBIND")]
        public void MergeDeleteAndBind()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2017/ENU/AutoCAD-NET/files/GUID-FAC1A5EB-2D9E-497B-8FD9-E11D2FF87B93-htm.html

            //https://adndevblog.typepad.com/autocad/2012/07/using-readdwgfile-with-net-attachxref-or-objectarx-acdbattachxref.html
            //Database oldDb = HostApplicationServices.WorkingDatabase; //is it necessary?

            

            // User should input the folder where the dwgs are saved
            PromptResult pr = ed.GetString("\nEnter folder containing DWGs to process: ");

            if (pr.Status != PromptStatus.OK)
                return;

            string pathName = pr.StringResult;

            string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

            // We'll use some counters to keep track
            // of how the processing is going
            int processed = 0, saved = 0, problem = 0;

            //var dict = File.ReadLines($"{pathName}\\summary.csv").Select(line => line.Split(',')).ToDictionary(line => line[0], line => line.ToList());

            //dict.Remove(dict.Keys.First()); //remove the csv header

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
                                //Point3d insPt = new Point3d(0, 0, sheetObject.viewCentre.z);
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

                            //forms.MessageBox.Show(lm.CurrentLayout);

                            string currentLo = lm.CurrentLayout;

                            DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                            Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;
                            
                            foreach (ObjectId ID in CurrentLo.GetViewports())
                            {
                                Viewport viewPort = trans.GetObject(ID, OpenMode.ForWrite) as Viewport;

                                //Point3d revitViewportCentre = new Point3d(double.Parse(dict[name][5]), double.Parse(dict[name][6]), 0);
                                XYZ vpCentre = sheetObject.viewportCentre;
                                Point3d revitViewportCentre = new Point3d(vpCentre.x, vpCentre.y, 0);

                                //Point3d revitViewCentreWCS = new Point3d(double.Parse(dict[name][1]), double.Parse(dict[name][2]), 0);
                                XYZ _revitViewCentreWCS = sheetObject.viewCentre;
                                //Point3d revitViewCentreWCS = new Point3d(revitViewCentre.x, revitViewCentre.y, 0);
                                Point3d revitViewCentreWCS = new Point3d(_revitViewCentreWCS.x, _revitViewCentreWCS.y, 0);

                                //double degrees = DegToRad(double.Parse(dict[name][4]));
                                double degrees = Helpers.DegToRad(sheetObject.angleToNorth);
                                //double vpWidht = double.Parse(dict[name][8]);
                                double vpWidht = sheetObject.viewportWidth;
                                //double vpHeight = double.Parse(dict[name][9]);
                                double vpHeight = sheetObject.viewportHeight;

                                Viewport viewPortToUpdate = null;


                                    if (viewPort != null && CurrentLo.GetViewports().Count == 2) //by default the Layout is a viewport too...https://forums.autodesk.com/t5/net/layouts-and-viewports/td-p/3128748
                                    {
                                        viewPortToUpdate = viewPort;
                                    }
                                    else if (viewPort != null && viewPort.CenterPoint.DistanceTo(revitViewportCentre) < 100)  //Should use the closest viewport, not a fixed distance
                                    {
                                        viewPortToUpdate = viewPort;
                                    }
                                    else
                                    {
                                        viewPort.FreezeLayersInViewport(layerToFreeze.GetEnumerator());
                                    }
                                    
                                Point3dCollection vpCorners = CadHelper.GetViewportBoundary(viewPortToUpdate);

                                Matrix3d mt = CadHelper.PaperToModel(viewPortToUpdate);

                                Point3dCollection vpCornersInModel = CadHelper.TransformPaperSpacePointToModelSpace(vpCorners, mt);

                                ObjectId[] viewportContent = CadHelper.SelectEntitisInModelSpaceByViewport(doc, vpCornersInModel);

                                if (viewportContent != null)
                                {
                                    foreach (ObjectId item in viewportContent)
                                    {
                                        Entity e = (Entity)trans.GetObject(item, OpenMode.ForWrite);
                                        e.Erase();
                                    }

                                }
                                else
                                {
                                    ed.WriteMessage("viewport content is null!");
                                }

                                Helpers.UpdateViewport(viewPortToUpdate, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);


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