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

            int counter = 0;

            try
            {
                using (var form = new Form1())
                {
                    //use ShowDialog to show the form as a modal dialog box. 
                    form.ShowDialog();

                    //if the user hits cancel just drop out of macro
                    if (form.DialogResult == winForm.DialogResult.Cancel)
                    {
                        return Result.Cancelled;
                    }

                    string destinationFolder = form.tBoxDestinationFolder;

                    string[] sheetNumbers = form.tBoxSheetNumber.Split(' ');

                    string exportSettings = form.tBoxExportSettings;

                    using (TransactionGroup tranGroup = new TransactionGroup(doc))
                    {

                        tranGroup.Start("Hide and Export");


                        using (Transaction t = new Transaction(doc))
                        {

                            foreach (string sheetNumber in sheetNumbers)
                            {

                                ViewSheet vs = allSheets.Where(x => x.SheetNumber == sheetNumber).First();

                                List<ElementId> views = vs.GetAllPlacedViews().ToList();

                                View planView = doc.GetElement(views.First()) as View;

                                ICollection<Element> fec = null;

                                t.Start("Hide elements");

                                if (planView.ViewType == ViewType.FloorPlan)
                                {
                                    fec = new FilteredElementCollector(doc, views.First())
                                                                                .WhereElementIsNotElementType()
                                                                                .ToElements()
                                                                                .Where(x => x.CanBeHidden(planView))
                                                                                .ToList();

                                    planView.HideElements(fec.Select(x => x.Id).ToList());
                                }

                                t.Commit();

                                //planView.ViewType == ViewType.EngineeringPlan
                                t.Start("Export Sheet");

                                if (!Helpers.ExportDWG(doc, vs, exportSettings, sheetNumber, destinationFolder))
                                {
                                    TaskDialog.Show("Error", "Check that the destination folder exists or the Export Settings exists");
                                }
                                else
                                {
                                    counter += 1;
                                }

                                t.Commit();

                                t.Start("Unhide elements");

                                if (null != fec)
                                    planView.UnhideElements(fec.Select(x => x.Id).ToList());

                                t.Commit();

                            }

                        }

                        tranGroup.Assimilate();
                    }

                }

                TaskDialog.Show("Done", $"{counter} sheets have been exported");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

        }
    }
}

