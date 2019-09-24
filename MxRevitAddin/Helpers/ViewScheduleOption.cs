using System.Collections.Generic;

namespace MxRevitAddin
{
    public class ViewScheduleOption
    {
        public string FormattedName
        {
            get
            {
                if (Name == "SHOW ALL SHEETS")
                {
                    return Name;
                }

                return string.Format("{0} ({1})", Name, ViewSheetCount);
            }

        }

        public string Name { get; set; }

        public int ViewSheetCount
        {
            get
            {
                int counter = 0;
                foreach (Autodesk.Revit.DB.View view in Views)
                {
                    if (view is Autodesk.Revit.DB.ViewSheet)
                    {
                        counter++;
                    }
                }
                return counter;
            }
        }

        public List<Autodesk.Revit.DB.ViewSheet> Views { get; set; }
    }
}
