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
    public class CommandExplodeBlockRef
    {
        //http://drive-cad-with-code.blogspot.com/2014/03/selecting-entities-in-modelspace.html

        [CommandMethod("EXPLODEBLOCKREF")]
        public void DeleteVP()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Autodesk.AutoCAD.DatabaseServices.TransactionManager bTransMan = doc.TransactionManager;

            LayoutManager lm = LayoutManager.Current;

            lm.CurrentLayout = "Model";

            Database db = doc.Database;

            //get document name
            ed.WriteMessage("\n======================== Xref Name: " + doc.Name + "\n");


            ed.Command("_.explode", "all", " ");

            ed.Command("_.zoom", "_extents");

            ed.WriteMessage("Save file \n");
            db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

            ed.WriteMessage("done");
        }

        //https://www.modelical.com/en/explode-autocad-block-references-using-net/
        public bool ExplodeBlockByNameCommand(string blockToExplode)
        {
            bool explodeResult = true;
            Document bDwg = Application.DocumentManager.MdiActiveDocument;
            Editor ed = bDwg.Editor;
            Database db = bDwg.Database;
            Autodesk.AutoCAD.DatabaseServices.TransactionManager bTransMan = bDwg.TransactionManager;
            using (Transaction bTrans = bTransMan.StartTransaction())
            {
                try
                {
                    BlockTable bt = (BlockTable)bTrans.GetObject(db.BlockTableId, OpenMode.ForRead);

                    ed.WriteMessage("\nProcessing {0}", blockToExplode);

                    if (bt.Has(blockToExplode))
                    {
                        ObjectId blkId = bt[blockToExplode];
                        BlockTableRecord btr = (BlockTableRecord)bTrans.GetObject(blkId, OpenMode.ForRead);
                        ObjectIdCollection blkRefs = btr.GetBlockReferenceIds(true, true);

                        foreach (ObjectId blkXId in blkRefs)
                        {
                            //create collection for exploded objects
                            DBObjectCollection objs = new DBObjectCollection();

                            //handle as entity and explode
                            Entity ent = (Entity)bTrans.GetObject(blkXId, OpenMode.ForRead);
                            ent.Explode(objs);
                            ed.WriteMessage("\nExploded an Instance of {0}", blockToExplode);

                            //erase Block
                            ent.UpgradeOpen();
                            ent.Erase();

                            BlockTableRecord btrCs = (BlockTableRecord)bTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                            foreach (DBObject obj in objs)
                            {
                                Entity ent2 = (Entity)obj;
                                btrCs.AppendEntity(ent2);
                                bTrans.AddNewlyCreatedDBObject(ent2, true);
                            }

                        }
                    }

                    bTrans.Commit();
                }
                catch
                {
                    ed.WriteMessage("\nSomething went wrong");
                    explodeResult = false;
                }
                finally
                {
                }
                ed.WriteMessage("\n");
                bTrans.Dispose();
                bTransMan.Dispose();
            }

            return explodeResult; //return wheter the method was succesful or not

        }

        //https://www.keanw.com/2014/09/exploding-nested-autocad-blocks-using-net.html

        private void ExplodeBlock(Transaction tr, Database db, ObjectId id, bool erase = true)

        {

            // Open out block reference - only needs to be readable for the explode operation, as it's non-destructive
            var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

            // We'll collect the BlockReferences created in a collection
            var toExplode = new ObjectIdCollection();

            // Define our handler to capture the nested block references
            ObjectEventHandler handler = (s, e) =>
            {
                if (e.DBObject is BlockReference)

                {

                    toExplode.Add(e.DBObject.ObjectId);

                }

            };


            // Add our handler around the explode call, removing it directly afterwards

            db.ObjectAppended += handler;

            br.ExplodeToOwnerSpace();

            db.ObjectAppended -= handler;



            // Go through the results and recurse, exploding the

            // contents



            foreach (ObjectId bid in toExplode)

            {

                ExplodeBlock(tr, db, bid, erase);

            }



            // We might also just let it drop out of scope



            toExplode.Clear();



            // To replicate the explode command, we're delete the

            // original entity



            if (erase)

            {

                br.UpgradeOpen();

                br.Erase();

                br.DowngradeOpen();

            }

        }
    }//close class
}//close namespace