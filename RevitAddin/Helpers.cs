using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAddin
{
    class Helpers
    {
        public static XYZ BBoxCenter(BoundingBoxXYZ bbox, Document doc, View vs)
        {

            XYZ min = bbox.Min;
            XYZ max = bbox.Max;
            XYZ center = (min + max) / 2;
            XYZ min0 = PointFlatten(min);
            XYZ max0 = PointFlatten(max);

            using (Transaction t = new Transaction(doc, "test"))
            {
                t.Start();

                //Line L1 = Line.CreateBound(min0, (min0+max0)/2);		
                //				Line L2 = Line.CreateBound(max0, (min0+max0)/2);

                //doc.Create.NewDetailCurve(vs,L1);
                //doc.Create.NewDetailCurve(vs,L2);


                t.Commit();
            }

            return (min0 + max0) / 2;
        }

        public static bool ExportDWG(Document document, Autodesk.Revit.DB.View view, string setupName, string xrefName, string folder)
        {
            bool exported = false;
            // Get the predefined setups and use the one with the given name.
            IList<string> setupNames = BaseExportOptions.GetPredefinedSetupNames(document);
            foreach (string name in setupNames)
            {
                if (name.CompareTo(setupName) == 0)
                {
                    // Export using the predefined options
                    DWGExportOptions dwgOptions = DWGExportOptions.GetPredefinedOptions(document, name);

                    // Export the active view
                    ICollection<ElementId> views = new List<ElementId>();
                    views.Add(view.Id);
                    // The document has to be saved already, therefore it has a valid PathName.
                    
                    exported = document.Export(folder, xrefName, views, dwgOptions);
                    break;
                }
            }

            return exported;
        }

        public static string PointToString(XYZ point)
        {

            string ptX = Convert.ToInt32(UnitUtils.ConvertFromInternalUnits(point.X, DisplayUnitType.DUT_MILLIMETERS)).ToString();
            string ptY = Convert.ToInt32(UnitUtils.ConvertFromInternalUnits(point.Y, DisplayUnitType.DUT_MILLIMETERS)).ToString();
            string ptZ = Convert.ToInt32(UnitUtils.ConvertFromInternalUnits(point.Z, DisplayUnitType.DUT_MILLIMETERS)).ToString();

            return String.Format("{0},{1},{2}", ptX, ptY, ptZ);
        }

        public static XYZ PointFlatten(XYZ point)
        {

            return new XYZ(point.X, point.Y, 0);
        }

        
    }
}
