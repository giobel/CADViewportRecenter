#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using winForm = System.Windows.Forms;

#endregion

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class ExportSheets : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            IEnumerable<ViewSheet> allSheets = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType().ToElements().Cast<ViewSheet>();

            //Handling and Dismissing a Warning Message
            //https://thebuildingcoder.typepad.com/blog/2013/03/export-wall-parts-individually-to-dxf.html
            uiapp.DialogBoxShowing += new EventHandler<DialogBoxShowingEventArgs>(OnDialogBoxShowing);

            IList<string> dWGExportOptions = DWGExportOptions.GetPredefinedSetupNames(doc);

            List<ViewScheduleOption> viewScheduleOptions = Helpers.GetViewScheduleOptions(doc);

            int counter = 0;

            try
            {
                using (var form = new Form1())
                {
                    //set the form export settings
                    form.cboxExportSettingsDataSource = dWGExportOptions;
                    //set the form sheets
                    form.cboxSheetDataSource = viewScheduleOptions;

                    //use ShowDialog to show the form as a modal dialog box. 
                    form.ShowDialog();

                    //if the user hits cancel just drop out of macro
                    if (form.DialogResult == winForm.DialogResult.Cancel)
                    {
                        return Result.Cancelled;
                    }

                    string destinationFolder = form.tBoxDestinationFolder;

                    //string[] sheetNumbers = form.tboxSelectedSheets.Split(' ');

                    string exportSettings = form.tBoxExportSettings;

                    DWGExportOptions dwgOptions = DWGExportOptions.GetPredefinedOptions(doc, form.tBoxExportSettings);

                    if (dwgOptions == null)
                    {
                        TaskDialog.Show("Error", "Export setting not found");
                        return Result.Failed;
                    }

                    if (dwgOptions.MergedViews == false)
                    {
                        TaskDialog.Show("Error", "Please unselect export view as external reference.");
                        return Result.Failed;
                    }

                    if (dwgOptions.TargetUnit != ExportUnit.Millimeter)
                    {
                        TaskDialog.Show("Error", "Export units not set to Millimeter. Please fix this before exporting.");
                        return Result.Failed;
                    }

                    ICollection<ElementId> categoryToIsolate = new List<ElementId>();

                    Categories groups = doc.Settings.Categories;

                    categoryToIsolate.Add(groups.get_Item(BuiltInCategory.OST_Loads).Id);

                    List<ViewSheet> selectedSheets = form.tboxSelectedSheets;

                    int n = form.tboxSelectedSheets.Count;
                    string s = "{0} of " + n.ToString() + " sheets exported...";
                    string caption = "Export Sheets";

                    using (ProgressForm pf = new ProgressForm(caption, s, n))
                    {

                        using (Transaction t = new Transaction(doc, "Hide categories"))
                        {
                            t.Start();

                            foreach (ViewSheet vs in selectedSheets)
                            {
                                if (pf.abortFlag)
                                    break;

                                //ViewSheet vs = allSheets.Where(x => x.SheetNumber == sheetNumber).First();
                                
                                string fileName = vs.LookupParameter("CADD File Name").AsString();

                                List<ElementId> views = vs.GetAllPlacedViews().ToList();

                                foreach (ElementId eid in views)
                                {
                                    View planView = doc.GetElement(eid) as View;

                                    if (planView.ViewType == ViewType.FloorPlan || planView.ViewType == ViewType.EngineeringPlan || planView.ViewType == ViewType.CeilingPlan)
                                    {
                                        planView.IsolateCategoriesTemporary(categoryToIsolate);
                                    }
                                }

                                if (!Helpers.ExportDWG(doc, vs, exportSettings, fileName, destinationFolder))
                                {
                                    TaskDialog.Show("Error", "Check that the destination folder exists");
                                }
                                else
                                {
                                    counter += 1;
                                }

                                pf.Increment();
                            }

                            t.RollBack();
                        }//close using transaction

                    }
                }//close using form

                TaskDialog.Show("Done", $"{counter} sheets have been exported");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }
            finally
            {
                uiapp.DialogBoxShowing -= new EventHandler<DialogBoxShowingEventArgs>(OnDialogBoxShowing);
            }

        }

        private void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs e)
        {
            TaskDialogShowingEventArgs e2 = e as TaskDialogShowingEventArgs;

            if (null != e2 && e2.DialogId.Equals(
              "TaskDialog_Really_Print_Or_Export_Temp_View_Modes"))
            {
                e.OverrideResult(
                  (int)TaskDialogResult.CommandLink2);
            }
        }
    }
}

