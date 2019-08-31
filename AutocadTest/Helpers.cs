using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AutocadTest
{
    class Helpers
    {

        //https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2016/ENU/AutoCAD-NET/files/GUID-288B4394-C51F-48CC-8B8C-A27873CFFDC1-htm.html
        public static void PurgeUnusedLayers(Transaction acTrans, Database acCurDb)
        {
            LayerTable acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;


            // Create an ObjectIdCollection to hold the object ids for each table record
            ObjectIdCollection acObjIdColl = new ObjectIdCollection();

            // Step through each layer and add iterator to the ObjectIdCollection
            foreach (ObjectId acObjId in acLyrTbl)
            {
                acObjIdColl.Add(acObjId);
            }

            // Remove the layers that are in use and return the ones that can be erased
            acCurDb.Purge(acObjIdColl);

            // Step through the returned ObjectIdCollection
            // and erase each unreferenced layer
            foreach (ObjectId acObjId in acObjIdColl)
            {
                SymbolTableRecord acSymTblRec;
                acSymTblRec = acTrans.GetObject(acObjId, OpenMode.ForWrite) as SymbolTableRecord;
                try
                {
                    // Erase the unreferenced layer
                    acSymTblRec.Erase(true);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception Ex)
                {
                    // Layer could not be deleted
                    //Application.ShowAlertDialog("Error:\n" + Ex.Message);
                }
            }
        }


        //https://www.keanw.com/2007/08/purging-registe.html
        public static int PurgeDatabase(Database db, Transaction tr)
        {
            int idCount = 0;
            
                // Create the list of objects to "purge"
                ObjectIdCollection idsToPurge = new ObjectIdCollection();
                
                // Add all the Registered Application names
                
                RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId,OpenMode.ForRead);


                foreach (ObjectId raId in rat)

                {

                    if (raId.IsValid)

                    {

                        idsToPurge.Add(raId);

                    }

                }


                // Call the Purge function to filter the list


                db.Purge(idsToPurge);


                // Erase each of the objects we've been

                // allowed to


                foreach (ObjectId id in idsToPurge)

                {

                    DBObject obj =

                      tr.GetObject(id, OpenMode.ForWrite);


                    // Let's just add to me "debug" code

                    // to list the registered applications

                    // we're erasing


                    RegAppTableRecord ratr =

                      obj as RegAppTableRecord;


                    obj.Erase();

                }


                // Return the number of objects erased

                // (i.e. purged)


                idCount = idsToPurge.Count;           

            return idCount;
    }

    public static void UpdateViewport(Viewport _vp, Point3d rvtCentre, Point3d rvtCentreWCS, double degrees, double vpWidth, double vpHeight)
        {
            _vp.UpgradeOpen();
            double cs = _vp.CustomScale; //save the original scale as it changes when we change viewport width and height
            _vp.CenterPoint = rvtCentre; //move the viewport to the revit location
            _vp.Width = vpWidth; //set the width to match the revit width
            _vp.Height = vpHeight; //idem
            _vp.CustomScale = cs;
            //_vp.BackClipOn = true;
            //_vp.FrontClipOn = true;
            //VP.Erase();

            TwistViewport(_vp.Id, rvtCentreWCS, degrees);
        }

        public static void BindXrefs(Database db)

        {
            ObjectIdCollection xrefCollection = new ObjectIdCollection();

            using (XrefGraph xg = db.GetHostDwgXrefGraph(false))

            {
                int numOfNodes = xg.NumNodes;

                for (int cnt = 0; cnt < xg.NumNodes; cnt++)

                {
                    XrefGraphNode xNode = xg.GetXrefNode(cnt) as XrefGraphNode;

                    if (!xNode.Database.Filename.Equals(db.Filename))

                    {

                        if (xNode.XrefStatus == XrefStatus.Resolved)

                        {

                            xrefCollection.Add(xNode.BlockTableRecordId);

                        }

                    }
                }
            }

            if (xrefCollection.Count != 0)

                db.BindXrefs(xrefCollection, true);
        }



        public static double DegToRad(double deg)
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

        public static void CreateLayer(Database acCurDb, Transaction acTrans, string sLayerName)
        {
            // Open the Layer table for read
            LayerTable layerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
            if (layerTable.Has(sLayerName) == false)
            {
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
            }

        }
   
    }
}
