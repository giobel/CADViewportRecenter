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
    public class DeleteInViewport
    {
        //http://drive-cad-with-code.blogspot.com/2014/03/selecting-entities-in-modelspace.html

        [CommandMethod("DELETEINVIEWPORT")]
        public void MergeAndBind()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            Database db = doc.Database;

            //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2017/ENU/AutoCAD-NET/files/GUID-FAC1A5EB-2D9E-497B-8FD9-E11D2FF87B93-htm.html

            //https://adndevblog.typepad.com/autocad/2012/07/using-readdwgfile-with-net-attachxref-or-objectarx-acdbattachxref.html

            string outputPath = @"C:\Users\giovanni.brogiolo\Desktop\MTCStandards";

            //using a Sheet Object
            var logFile = File.ReadAllLines(@"C:\Users\giovanni.brogiolo\Desktop\MTCStandards\summary.csv").Select(line => line.Split(',')).ToList<string[]>();
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
                    LayoutManager lm = LayoutManager.Current;

                    lm.CurrentLayout = "Model"; //is it necessary?

                try
                {


                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {

                        lm.CurrentLayout = "Layout1";

                        ed.Command("_.zoom", "_extents");

                        //forms.MessageBox.Show(lm.CurrentLayout);

                        string currentLo = lm.CurrentLayout;

                        DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                        Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;

                        foreach (ObjectId ID in CurrentLo.GetViewports())
                        {
                            Viewport VP = trans.GetObject(ID, OpenMode.ForWrite) as Viewport;

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


                            if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre) < 100)  //Should use the closest viewport, not a fixed distance
                            {
                                Point3dCollection vpBoundaryPts = CadHelper.GetViewportBoundary(VP);

                                LayoutManager.Current.CurrentLayout = "Model";

                                ObjectId[] ents = MyCommands.SelectEntitisInModelSpaceByViewport(doc, vpBoundaryPts);

                                foreach (ObjectId item in ents)
                                {
                                    Entity e = (Entity)trans.GetObject(item, OpenMode.ForWrite);
                                    e.Erase();
                                }

                                LayoutManager.Current.CurrentLayout = "Layout1";

                                Helpers.UpdateViewport(VP, revitViewportCentre, revitViewCentreWCS, degrees, vpWidht, vpHeight);
                            }
                            else
                            {
                                
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
                }
                
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("\nProblem processing file: {0} - \"{1}\"", sheetObject.sheetName, ex.Message);

                    }
                }

            ed.WriteMessage("done");
        }





    }
}