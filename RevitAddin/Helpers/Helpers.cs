﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        public static bool ExportDWG(Document document, Autodesk.Revit.DB.View view, string setupName, string fileName, string folder)
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
                    exported = document.Export(folder, fileName, views, dwgOptions);
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

        /// <summary>
        /// Retrieves all View Schedule Options in the Document. Copy paste from Revup
        /// </summary>
        public static List<ViewScheduleOption> GetViewScheduleOptions(Document doc)
        {
            
            List<ViewScheduleOption> options = new List<ViewScheduleOption>();

            var collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
            if (collector.Any())
            {
                var allNonCategorySchedules = collector.ToElements().Cast<ViewSchedule>().ToList().FindAll(vs => !vs.IsTemplate && vs.Definition.CategoryId == new ElementId(BuiltInCategory.OST_Sheets));
                if (allNonCategorySchedules.Any())
                {
                    foreach (var schedule in allNonCategorySchedules)
                    {
                        collector = new FilteredElementCollector(doc, schedule.Id).OfClass(typeof(ViewSheet));
                        if (collector.Any())
                        {
                            options.Add(new ViewScheduleOption() { Name = schedule.ViewName, Views = collector.ToElements().Cast<ViewSheet>().ToList() });
                        }
                    }
                }
            }

            return options;
        }

        public static List<Outline> GetOutline(Document doc, List<View> viewList)
        {
            List<Outline> outlines = new List<Outline>();

            foreach (View pView in viewList)
            {
                BoundingBoxXYZ bbox = pView.get_BoundingBox(pView);
                Outline currentOutline = new Outline(bbox.Min, bbox.Max);
                outlines.Add(currentOutline);
            }
            return outlines;
        }

        /// <summary>
        /// Check whether a viewport outline intersects a list of other Viewport outlins
        /// </summary>
        /// <param name="current">The outline to check</param>
        /// <param name="check">The list of outlines to be checked against</param>
        /// <returns></returns>
        public static bool OutlineIntersects(Outline current, List<Outline> check)
        {
            foreach (Outline element in check)
            {
                if (current.Intersects(element, 0))
                {
                    return true;
                }
            }
            return false;
        }

        public static string CheckVPOverlaps(Document doc, List<View> views)
        {
            List<Outline> outlineList = GetOutline(doc, views);

            StringBuilder sb = new StringBuilder();

            Outline current = null;

            int counter = 0;

            while (counter < views.Count)
            {
                current = outlineList.First();
                outlineList.Remove(current);

                if (OutlineIntersects(current, outlineList))
                {
                    sb.AppendLine(views.ElementAt(counter).ViewName+ " xref");
                }
                else
                {
                    sb.AppendLine(views.ElementAt(counter).ViewName + " bind");
                }
                counter += 1;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Loop through all the views placed on a sheet and return a list of plan views only (Floor, Ceiling, Engineering or Area Plan only).
        /// </summary>
        /// <param name="planViews"></param>
        /// <returns></returns>
        public static List<View> FilterPlanViewport(Document doc, ISet<ElementId> planViewsIds)
        {
            List<View> filteredViews = new List<View>();

            foreach (ElementId eid in planViewsIds)
            {
                View planView = doc.GetElement(eid) as View;
                if (planView.ViewType == ViewType.FloorPlan || planView.ViewType == ViewType.EngineeringPlan || planView.ViewType == ViewType.CeilingPlan || planView.ViewType == ViewType.AreaPlan)
                {
                    filteredViews.Add(planView);
                }

            }
            return filteredViews;
        }
    }
}