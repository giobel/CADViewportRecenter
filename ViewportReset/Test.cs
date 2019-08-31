using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using System;
using forms = System.Windows.Forms;
using System.Linq;
using Autodesk.AutoCAD.Geometry;
using System.Reflection;
using Autodesk.AutoCAD.PlottingServices;
using System.Collections.Generic;
using Autodesk.AutoCAD.Colors;
using ViewportReset;

namespace ViewportReset
{
    public class Test
    {
        //http://drive-cad-with-code.blogspot.com/2014/03/selecting-entities-in-modelspace.html

        [CommandMethod("TEST")]
        public void Testa()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Layout1"; 

            ed.Command("_.zoom", "_extents");

            Database db = doc.Database;



            object obj = Application.GetSystemVariable("DBMOD");

            // Check the value of DBMOD, if 0 then the drawing has no unsaved changes
            if (System.Convert.ToInt16(obj) != 0)
            {
                
                    
                    db.SaveAs(doc.Name, true, DwgVersion.Current,doc.Database.SecurityParameters);
                
            }


            ed.WriteMessage("done");
        }


    }
}