using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using winForm = System.Windows.Forms;

namespace MxRevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class Help : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {

            TaskDialog myDialog = new TaskDialog("Help");
            myDialog.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            myDialog.MainContent = $"Before proceeding to export please check that these parameters exist:\n" +
                $"\n1. 'CADD File Name' applied to Sheets\n" +
                $"\n2. 'Mx Export_Sheet Filter' applied to Sheets\n" +
                $"\n3. 'Mx Keyplan (Y as required)' applied to Views. Views marked with Y will not be exported as xref.\n" +
                $"\nThen run in sequence:\n" +
                $"\n1. Check Sheets: groups the sheets in No Plan Views, Plans not Overlapping and Plans Overlapping. It populates the Mx Export_Sheet Filter with one of these values.\n" +
                $"\n2. Check Vieport Size: outputs the viewports sizes larger than the max and min values specified.\n" +
                $"\n3. Export Sheets: exports the sheets and renames per CAAD File Name parameter. User can choose to hide the content of the viewports " +
                $"containing a floor/ceiling/structural/area plan\n" +
                $"\n4. Export Xrefs+CSV: export the viewports containing a floor/ceiling/structural/area plan and creates a csv file to be read by the Autocad script.\n" +
                $"\nFor assistance please contact: Harry Suh, Keatan Howards or Tristan Oakley.";
            
            TaskDialogResult res = myDialog.Show();

            return Result.Succeeded;
        }//close execute
    }
}
