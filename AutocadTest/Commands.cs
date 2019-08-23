using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System.IO;

namespace AutocadTest
{
    public class Commands
    {


        [CommandMethod("CreateAndAssignALayer")]
        public static void CreateAndAssignALayer()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                                OpenMode.ForRead) as LayerTable;

                string sLayerName = "Center";

                if (acLyrTbl.Has(sLayerName) == false)
                {
                    using (LayerTableRecord acLyrTblRec = new LayerTableRecord())
                    {
                        // Assign the layer the ACI color 3 and a name
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                        acLyrTblRec.Name = sLayerName;

                        // Upgrade the Layer table for write
                        acLyrTbl.UpgradeOpen();

                        // Append the new layer to the Layer table and the transaction
                        acLyrTbl.Add(acLyrTblRec);
                        acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                    }
                }

                // Open the Block table for read
                BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord blockTableRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                //string PathName = $"{pathName}\\{dict[name][10]}";
                string PathName = @"C:\Users\giovanni.brogiolo\Desktop\MTCStandards\E2200-7163394-xref";
                //ObjectId acXrefId = db.AttachXref(PathName, dict[name][10]);
                ObjectId acXrefId = acCurDb.AttachXref(PathName, "xref");

                if (!acXrefId.IsNull)
                {
                    // Attach the DWG reference to the current space
                    //Point3d insPt = new Point3d(0, 0, sheetObject.viewCentre.z);
                    Point3d insPt = new Point3d(0, 0, 0);
                    using (BlockReference blockRef = new BlockReference(insPt, acXrefId))
                    {
                        blockRef.Layer = sLayerName;
                        BlockTable blocktable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = acTrans.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        modelSpace.AppendEntity(blockRef);

                        acTrans.AddNewlyCreatedDBObject(blockRef, true);
                    }
                }

                // Save the changes and dispose of the transaction
                acTrans.Commit();
            }//close transaction
            acDoc.Editor.Command("_.ZOOM", "e");



        }//close command


        [CommandMethod("OpenAndSave")]
        public static void OpenAndSaveDoc()
        {
            Database acCurDb = new Database(false, false);
            string filePath = @"C:\Users\giovanni.brogiolo\Desktop\MTCStandards";
            string fileName = "E2200";


            acCurDb.ReadDwgFile($"{filePath}\\{fileName}.dwg", FileShare.ReadWrite, true, "");
            //acCurDb.CloseInput(true);

            string sLayerName = "myLayero";

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                
                // Open the Layer table for read
                LayerTable layerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                
                    using (LayerTableRecord acLyrTblRec = new LayerTableRecord())
                    {
                        // Assign the layer the ACI color 3 and a name
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                        acLyrTblRec.Name = sLayerName;

                        // Upgrade the Layer table for write
                        layerTable.UpgradeOpen();

                        // Append the new layer to the Layer table and the transaction
                        layerTable.Add(acLyrTblRec);
                        acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                    }

                ObjectId layer = new ObjectId() ;
                foreach (ObjectId layerId in layerTable)
                {
                    LayerTableRecord currentLayer = acTrans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                    if (currentLayer.Name == sLayerName)
                        layer = layerId;
                }


                    string PathName = $"{filePath}\\E2200-7163394-xref.dwg";
                    
                    ObjectId acXrefId = acCurDb.AttachXref(PathName, "ciao");


                    if (!acXrefId.IsNull)
                    {
                        // Attach the DWG reference to the current space
                        //Point3d insPt = new Point3d(0, 0, sheetObject.viewCentre.z);
                        Point3d insPt = new Point3d(0, 0, 0);
                        using (BlockReference blockRef = new BlockReference(insPt, acXrefId))
                        {

                            blockRef.SetLayerId(layer,true);
                            BlockTable blocktable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                            BlockTableRecord modelSpace = acTrans.GetObject(blocktable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                            modelSpace.AppendEntity(blockRef);

                            acTrans.AddNewlyCreatedDBObject(blockRef, true);
                        }
                    }


                

                acTrans.Commit();
            }

            acCurDb.SaveAs($"{filePath}\\{fileName}-changed.dwg",DwgVersion.Current);
        }
    }
}
