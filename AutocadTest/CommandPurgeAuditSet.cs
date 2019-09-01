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

namespace TristanBatchCommands
{
    public class CommandPurgeAuditSet
    {
        [CommandMethod("PURGEAUDITSET")]
        public void CreateLayers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            LayoutManager lm = LayoutManager.Current;

            ed.WriteMessage("Swith to Model layout \n");
            lm.CurrentLayout = "Model";

            ed.Command("_.zoom", "_extents");

            //ed.WriteMessage("======================== Run Set by layer\n");
            //ed.Command("-setbylayer", "all");

            ed.WriteMessage("======================== Run Audit\n");
            ed.Command("audit", "y");

            ed.WriteMessage("======================== Run Purge \n");
            ed.Command("-purge", "all", " ", "n");

            ed.WriteMessage("Save file \n");
            db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

            lm.CurrentLayout = "Layout1";

            ed.WriteMessage("done");
        }


    }//close class
}//close namespace