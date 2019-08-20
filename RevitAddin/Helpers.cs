using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

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

            //using (Transaction t = new Transaction(doc, "test"))
            //{
            //    t.Start();

            //    //Line L1 = Line.CreateBound(min0, (min0+max0)/2);		
            //    //				Line L2 = Line.CreateBound(max0, (min0+max0)/2);

            //    //doc.Create.NewDetailCurve(vs,L1);
            //    //doc.Create.NewDetailCurve(vs,L2);


            //    t.Commit();
            //}

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

            return String.Format("{0},{1},0", ptX, ptY);
        }

        public static XYZ PointFlatten(XYZ point)
        {

            return new XYZ(point.X, point.Y, 0);
        }

        /// <summary>
        /// https://stackoverflow.com/questions/9815699/how-to-calculate-centroid
        /// http://coding-experiments.blogspot.com/2009/09/xna-quest-for-centroid-of-polygon.html
        /// </summary>
        /// <param name="nodes"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public static XYZ GetCentroid(List<XYZ> nodes, int count)
        {
            double x = 0, y = 0, area = 0, k;
            XYZ a, b = nodes[count - 1];

            for (int i = 0; i < count; i++)
            {
                a = nodes[i];

                k = a.Y * b.X - a.X * b.Y;
                area += k;
                x += (a.X + b.X) * k;
                y += (a.Y + b.Y) * k;

                b = a;
            }
            area *= 3;

            return (area == 0) ? XYZ.Zero : new XYZ(x /= area, y /= area, 0);
        }

    }
}
