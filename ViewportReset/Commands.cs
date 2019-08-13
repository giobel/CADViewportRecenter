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
        [CommandMethod("UAIF")]
        public void UpdateAttributeInFiles()
        {
            Document doc =
              Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Have the user choose the block and attribute
            // names, and the new attribute value
            PromptResult pr =ed.GetString("\nEnter folder containing DWGs to process: ");

            if (pr.Status != PromptStatus.OK)
                return;

            string pathName = pr.StringResult;

            string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

            // We'll use some counters to keep track
            // of how the processing is going
            int processed = 0, saved = 0, problem = 0;

            foreach (string fileName in fileNames)
            {
                if (fileName.EndsWith(".dwg",StringComparison.CurrentCultureIgnoreCase))
                {
                    string outputName =fileName.Substring(0,fileName.Length - 4) +"_updated.dwg";

                    Database db = new Database(false, false);
                    using (db)
                    {
                        try
                        {
                            ed.WriteMessage("\n\nProcessing file: " + fileName);

                            db.ReadDwgFile(fileName,FileShare.ReadWrite,false,"");

                            //int attributesChanged = UpdateAttributesInDatabase(
                            //    db,
                            //    blockName,
                            //    attbName,
                            //    attbValue
                            //  );

                            
                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                LayoutManager lm = LayoutManager.Current;
                                
                                lm.CurrentLayout = "Layout1";

                                //forms.MessageBox.Show(lm.CurrentLayout);

                                string currentLo = lm.CurrentLayout;

                                DBDictionary LayoutDict = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                                Layout CurrentLo = trans.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;



                                //BlockTableRecord BlkTblRec = trans.GetObject(CurrentLo.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                                foreach (ObjectId ID in CurrentLo.GetViewports())
                                {
                                    Viewport VP = trans.GetObject(ID, OpenMode.ForRead) as Viewport;

                                    

                                    if (VP != null && VP.CenterPoint.DistanceTo(new Point3d(381,285,0))<10)
                                    {
                                        VP.UpgradeOpen();
                                        VP.Erase();

                                        //TwistViewport(VP.Id,new Point3d(75045.0075828448, 13265.0382509219, 0), DegToRad(-36));
                                    }
                                }

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


        [CommandMethod("UA")]
        public void UpdateAttribute()
        {
            Document doc =
              Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            // Have the user choose the block and attribute
            // names, and the new attribute value
            PromptResult pr =
              ed.GetString(
                "\nEnter name of block to search for: "
              );
            if (pr.Status != PromptStatus.OK)
                return;
            string blockName = pr.StringResult.ToUpper();
            pr =
              ed.GetString(
                "\nEnter tag of attribute to update: "
              );
            if (pr.Status != PromptStatus.OK)
                return;
            string attbName = pr.StringResult.ToUpper();
            pr =
              ed.GetString(
                "\nEnter new value for attribute: "
              );
            if (pr.Status != PromptStatus.OK)
                return;
            string attbValue = pr.StringResult;
            ed.WriteMessage(
              "\nProcessing file: " + db.Filename
            );
            int count =
              UpdateAttributesInDatabase(
                db,
                blockName,
                attbName,
                attbValue
              );
            ed.Regen();
            // Display the results
            ed.WriteMessage(
              "\nUpdated {0} instance{1} of " +
              "attribute {2}.",
              count,
              count == 1 ? "" : "s",
              attbName
            );
        }
        private int UpdateAttributesInDatabase(Database db,string blockName,string attbName,string attbValue)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Get the IDs of the spaces we want to process
            // and simply call a function to process each
            ObjectId msId, psId;

            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);

                msId = bt[BlockTableRecord.ModelSpace];

                psId = bt[BlockTableRecord.PaperSpace];
                // Not needed, but quicker than aborting
                tr.Commit();
            }

            int msCount =
              UpdateAttributesInBlock(
                msId,
                blockName,
                attbName,
                attbValue
              );
            int psCount =
              UpdateAttributesInBlock(
                psId,
                blockName,
                attbName,
                attbValue
              );
            return msCount + psCount;
        }
        private int UpdateAttributesInBlock(ObjectId btrId,string blockName,string attbName,string attbValue)
        {
            // Will return the number of attributes modified
            int changedCount = 0;
            Document doc =
              Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Transaction tr =
              doc.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTableRecord btr =
                  (BlockTableRecord)tr.GetObject(
                    btrId,
                    OpenMode.ForRead
                  );
                // Test each entity in the container...
                foreach (ObjectId entId in btr)
                {
                    Entity ent =
                      tr.GetObject(entId, OpenMode.ForRead)
                      as Entity;
                    if (ent != null)
                    {
                        BlockReference br = ent as BlockReference;
                        if (br != null)
                        {
                            BlockTableRecord bd =
                              (BlockTableRecord)tr.GetObject(
                                br.BlockTableRecord,
                                OpenMode.ForRead
                            );
                            // ... to see whether it's a block with
                            // the name we're after
                            if (bd.Name.ToUpper() == blockName)
                            {
                                // Check each of the attributes...
                                foreach (
                                  ObjectId arId in br.AttributeCollection
                                )
                                {
                                    DBObject obj =
                                      tr.GetObject(
                                        arId,
                                        OpenMode.ForRead
                                      );
                                    AttributeReference ar =
                                      obj as AttributeReference;
                                    if (ar != null)
                                    {
                                        // ... to see whether it has
                                        // the tag we're after
                                        if (ar.Tag.ToUpper() == attbName)
                                        {
                                            // If so, update the value
                                            // and increment the counter
                                            ar.UpgradeOpen();
                                            ar.TextString = attbValue;
                                            ar.DowngradeOpen();
                                            changedCount++;
                                        }
                                    }
                                }
                            }
                            // Recurse for nested blocks
                            changedCount +=
                              UpdateAttributesInBlock(
                                br.BlockTableRecord,
                                blockName,
                                attbName,
                                attbValue
                              );
                        }
                    }
                }
                tr.Commit();
            }
            return changedCount;
        }
    }
}