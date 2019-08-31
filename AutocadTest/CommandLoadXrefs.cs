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

namespace AutocadTest
{
    public class CommandLoadXrefs
    {
        //http://drive-cad-with-code.blogspot.com/2014/03/selecting-entities-in-modelspace.html

        [CommandMethod("LOADXREFS")]
        public void LoadXrefs()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Layout1";

            Database db = doc.Database;

            string dwgName = Path.GetFileNameWithoutExtension(doc.Name);

            string folderPath = Path.GetDirectoryName(doc.Name);

            List<SheetObject> sheetObjects = Helpers.SheetsObjectsFromCSV(folderPath, dwgName);

            List<ObjectId[]> viewportContentList = new List<ObjectId[]>();
            //get document name
            ed.WriteMessage("\n======================== Dwg Name: " + doc.Name + "\n");


            foreach (SheetObject sheetObject in sheetObjects)
            {
                ed.WriteMessage("======================== Xref(s): " + sheetObject.xrefName + "\n");

                string layerName = $"0-{sheetObject.xrefName}";

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
                    trans.Commit();
                }//close transaction
            }

            ed.WriteMessage($"======================== Check xref group for binding\n");
            if (sheetObjects.First().group == "bind")
            {
                Helpers.BindXrefs(db);
                ed.WriteMessage($"======================== xrefs binded\n");
            }

            ed.WriteMessage("Save file \n");
            db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

            ed.WriteMessage("done");
        }

    }//close class
}//close namespace