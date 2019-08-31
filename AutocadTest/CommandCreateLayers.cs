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
    public class CommandCreateLayers
    {
        [CommandMethod("CREATELAYERS")]
        public void CreateLayers()
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


                using (Transaction trans = db.TransactionManager.StartTransaction())
                {

                    string layerName = $"0-{sheetObject.xrefName}";

                    Helpers.CreateLayer(db, trans, layerName);

                    ed.WriteMessage("======================== Layer created: " + layerName + "\n");

                    trans.Commit();
                }//close transaction

            }

            ed.WriteMessage("Save file \n");
            db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

            ed.WriteMessage("done");
        }


    }//close class
}//close namespace