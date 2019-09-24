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
    public class CommandLoadAndBindXrefs
    {

        [CommandMethod("LOADANDBINDXREFS")]
        public void LoadXrefs()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== +++ Command Load and Bind Xrefs Start +++");

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Layout1";

            Database db = doc.Database;

            try
            {

                string dwgName = Path.GetFileNameWithoutExtension(doc.Name);

                string folderPath = Path.GetDirectoryName(doc.Name);

                List<SheetObject> sheetObjects = Helpers.SheetsObjectsFromCSV(folderPath, dwgName);

                List<ObjectId[]> viewportContentList = new List<ObjectId[]>();
                //get document name
                ed.WriteMessage("\n=== Dwg Name: " + doc.Name + "\n");

                //to bind after have been loaded
                List<string> xrefsToBind = new List<string>();

                foreach (SheetObject sheetObject in sheetObjects)
                {
                    ed.WriteMessage("=== Xref(s): " + sheetObject.xrefName + "\n");

                    string layerName = $"0-{sheetObject.xrefName}";

                    //redundant. We want to bind all the xrefs
                    //if (sheetObject.group == "bind")
                    //{
                        xrefsToBind.Add(sheetObject.xrefName);
                    //}

                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {

                        LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                        ObjectId layer = new ObjectId();

                        foreach (ObjectId layerId in layerTable)
                        {
                            LayerTableRecord currentLayer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                            if (currentLayer.Name == layerName)
                            {
                                layer = layerId;

                            }
                        }

                        //Load Xref
                        //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2015/ENU/AutoCAD-NET/files/GUID-D6EE5FE7-C0BC-4E9B-AAE3-3B5A14B870FE-htm.html
                        #region
                        string PathName = $"{folderPath}\\{sheetObject.xrefName}";

                        ObjectId acXrefId = db.AttachXref(PathName, sheetObject.xrefName);

                        if (!acXrefId.IsNull)
                        {
                            // Attach the DWG reference to the model space
                            Point3d insPt = new Point3d(0, 0, 0);
                            using (BlockReference blockRef = new BlockReference(insPt, acXrefId))
                            {
                                //load the xref to Layer 0
                                //blockRef.SetLayerId(layer, true);
                                BlockTable blocktable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                BlockTableRecord modelSpace = trans.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                modelSpace.AppendEntity(blockRef);

                                trans.AddNewlyCreatedDBObject(blockRef, true);
                            }

                            ed.WriteMessage("=== xref loaded\n");
                        }

                        #endregion


                        trans.Commit();
                    }//close transaction


                }

                ed.WriteMessage($"=== Check if the xref has to be bound\n");
                
                if (xrefsToBind.Count > 0)
                {
                    foreach (string xrefName in xrefsToBind)
                    {
                        //it seems nested xrefs are moved back to origin when binded...
                        //Helpers.BindXrefs(db);
                        //try this:
                        ed.Command("-xref", "bind", xrefName, " ");
                        ed.WriteMessage($"=== xrefs binded\n");
                    }
                }
                
                ed.WriteMessage($"=== Check xref group for binding\n");
                
                ed.WriteMessage("Save file \n");

                db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

                ed.WriteMessage("\n=== +++ Command Load and Bind Xrefs End +++");
            }//close try
            catch { }
            finally
            {
            }
        }

    }//close class
}//close namespace