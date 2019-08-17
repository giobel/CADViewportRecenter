using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using System;
using forms = System.Windows.Forms;
using System.Linq;
using Autodesk.AutoCAD.Geometry;

namespace AttributeUpdater
{
    public class Commands
    {
        [CommandMethod("TRISTAN")]
        public void UpdateAttributeInFiles()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            
            // User should input the folder where the dwgs are saved
            PromptResult pr =ed.GetString("\nEnter folder containing DWGs to process: ");

            if (pr.Status != PromptStatus.OK)
                return;

            string pathName = pr.StringResult;

            string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

            // We'll use some counters to keep track
            // of how the processing is going
            int processed = 0, saved = 0, problem = 0;

            var dict = File.ReadLines($"{pathName}\\summary.csv").Select(line => line.Split(',')).ToDictionary(line => line[0], line => line.ToList());

            dict.Remove(dict.Keys.First()); //remove the csv header

            foreach (string fileName in fileNames)
            {
                if (fileName.EndsWith(".dwg",StringComparison.CurrentCultureIgnoreCase))
                {
                    FileInfo fi = new FileInfo(fileName);

                    string name = fi.Name.Substring(0,fi.Name.Length-4);
                    string outputName = fileName.Substring(0, fileName.Length - 4) + "_updated.dwg";

                    Database db = new Database(false, false);
                    using (db)
                    {
                        try
                        {
                            ed.WriteMessage("\n\nProcessing file: " + fileName);

                            db.ReadDwgFile(fileName,FileShare.ReadWrite,false,"");

                            LayoutManager lm = LayoutManager.Current;

                            lm.CurrentLayout = "Model";

                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                //Attch Xref
                                string PathName = $"{pathName}\\{dict[name][10]}";
                                ObjectId acXrefId = db.AttachXref(PathName, dict[name][10]);

                                

                                if (!acXrefId.IsNull)
                                {
                                    // Attach the DWG reference to the current space
                                    Point3d insPt = new Point3d(0, 0, 0);
                                    using (BlockReference acBlkRef = new BlockReference(insPt, acXrefId))
                                    {
                                        BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                        BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                        modelSpace.AppendEntity(acBlkRef);

                                        //BlockTableRecord acBlkTblRec;
                                        //acBlkTblRec = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                        //acBlkTblRec.AppendEntity(acBlkRef);

                                        trans.AddNewlyCreatedDBObject(acBlkRef, true);

                                        
                                    }
                                }

                                using (ObjectIdCollection acXrefIdCol = new ObjectIdCollection())
                                {
                                    acXrefIdCol.Add(acXrefId);

                                    db.BindXrefs(acXrefIdCol, false);
                                }


                                lm.CurrentLayout = "Layout1";

                                //forms.MessageBox.Show(lm.CurrentLayout);

                                string currentLo = lm.CurrentLayout;

                                DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                                Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;



                                //BlockTableRecord BlkTblRec = trans.GetObject(CurrentLo.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                                foreach (ObjectId ID in CurrentLo.GetViewports())
                                {
                                    Viewport VP = trans.GetObject(ID, OpenMode.ForRead) as Viewport;


                                    Point3d revitViewportCentre = new Point3d(double.Parse(dict[name][5]), double.Parse(dict[name][6]), 0); 

                                    Point3d revitViewCentreWCS = new Point3d(double.Parse(dict[name][1]), double.Parse(dict[name][2]), 0);

                                    if (VP != null && VP.CenterPoint.DistanceTo(revitViewportCentre)<100)  //Should use the closest viewport, not a fixed distance!
                                    {
                                        VP.UpgradeOpen();
                                        double cs = VP.CustomScale; //save the original scale as it changes when we change viewport width and height
                                        VP.CenterPoint = revitViewportCentre; //move the viewport to the revit location
                                        VP.Width = double.Parse(dict[name][8]); //set the width to match the revit width
                                        VP.Height = double.Parse(dict[name][9]); //idem
                                        VP.CustomScale = cs;
                                        //VP.Erase();

                                        TwistViewport(VP.Id,revitViewCentreWCS, DegToRad(double.Parse(dict[name][4])));

                                
                                    }
                                }

                                //ed.Command("_.ZOOM", "_E");

                                //ed.Command("_.ZOOM", ".7X");

                                //ed.Regen();

                                trans.Commit();
                            }

                                ed.WriteMessage("\nSaving to file: {0}", outputName);

                                db.SaveAs(outputName,DwgVersion.Current);

                                saved++;
                            
                            processed++;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nProblem processing file: {0} - \"{1}\"",fileName,ex.Message);

                            problem++;
                        }
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


        static double DegToRad(double deg)
        {
            return deg * Math.PI / 180.0;

        }

        private static void TwistViewport(ObjectId vpId, Point3d target, double angle)
        {


            using (Transaction tran = vpId.Database.TransactionManager.StartTransaction())
            {
                Viewport vport = (Viewport)tran.GetObject(vpId, OpenMode.ForWrite);
                vport.Locked = false;
                vport.ViewDirection = Vector3d.ZAxis;
                vport.ViewTarget = target;
                vport.ViewCenter = Point2d.Origin;
                vport.TwistAngle = Math.PI * 2 - angle;
                //vport.Locked = true;

                tran.Commit();
            }
        }

    }
}