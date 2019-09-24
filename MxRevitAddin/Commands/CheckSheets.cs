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
    public class CheckSheets : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {

            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Application app = uiapp.Application;
                Document doc = uidoc.Document;

                string outputFile = @"C:\Temp\reportSheetSummary.csv";

                StringBuilder sb = new StringBuilder();

                //store the name of the missing parameter (is any). Either 'Mx Export_Sheet Filter' or 'CAAD File Name'
                string paramError = "";

                string paramExportSheetFilter = "Mx Export_Sheet Filter";

                try
                {
                    File.WriteAllText(outputFile,
                      "Sheet Number," +
                      "Group" +
                      Environment.NewLine
                     );

                    //List<ViewScheduleOption> viewScheduleOptions = Helpers.GetViewScheduleOptions(doc);

                    List<ViewSchedule> viewSchedule = Helpers.SheetList(doc);

                    
                    using (var form = new SheetSummaryForm())
                    {
                        //set the form sheets
                        form.cboxSheetDataSource = viewSchedule;

                        //set the form title
                        form.Text = "Populate Sheet Filter";

                        //use ShowDialog to show the form as a modal dialog box. 
                        form.ShowDialog();

                        //if the user hits cancel just drop out of macro
                        if (form.DialogResult == winForm.DialogResult.Cancel)
                        {
                            return Result.Cancelled;
                        }

                        //selected Sheet List
                        ViewSchedule selectedSheetList = form.selectedViewSchedule;

                        //Sheets in selected sheet list
                        var selectedSheets = new FilteredElementCollector(doc, selectedSheetList.Id).OfClass(typeof(ViewSheet)).ToElements().Cast<ViewSheet>();

                        int n = selectedSheets.Count();
                        string s = "{0} of " + n.ToString() + " sheets processed...";
                        string caption = "Find oversized views";

                        int countOverlappingViewports = 0;
                        int countNonOverlappingViewports = 0;
                        int countNoPlansViewports = 0;
                        int keynotesCount = 0;

                        string sheetGroup = ""; //No plans, Plans not Overlapping, Plans Overlapping, Keynotes, No Keynotes

                        //Find the sheets that have plans with keynotes on it. Note: this does not ensure there is a keynote on the Sheet! Could not find
                        //a better way of doing this
                        string sheetsWithKeynotes = Helpers.FindSheetsWithKeynotesOnPlan(doc);

                        using (Transaction t = new Transaction(doc, "Set Sheet Parameter"))
                        {
                            t.Start();

                            using (ProgressForm pf = new ProgressForm(caption, s, n))
                            {
                                foreach (ViewSheet vs in selectedSheets)
                                {
                                    if (pf.abortFlag)
                                        break;

                                    //if Mx Export_Sheet Filter param does not exists, exit
                                    if (vs.LookupParameter(paramExportSheetFilter) == null)
                                    {
                                        paramError = $"{paramExportSheetFilter} not found on Sheets.";
                                        throw new System.NullReferenceException();
                                    }

                                    //select all viewports on the current sheet..not needed TBC
                                    ICollection <ElementId> viewports = vs.GetAllViewports();

                                    //Select all the views on a sheet
                                    ISet<ElementId> placedViewsIds = vs.GetAllPlacedViews();

                                    //if "Mx Keyplan (Y as required) param does not exists, exit
                                    if (placedViewsIds.Count > 0)
                                    {
                                        if (doc.GetElement(placedViewsIds.First()).LookupParameter("Mx Keyplan (Y as required)") == null)
                                        {
                                            paramError = "'Mx Keyplan (Y as required)' not found on Views.";
                                            throw new System.NullReferenceException();
                                        }
                                    }


                                    //Filter the views that are plans (Floor,Ceiling,Engineering or Area)
                                    List<View> planViews = Helpers.FilterPlanViewport(doc, placedViewsIds);


                                    if (planViews.Count > 0) //sheet has plan views
                                    {

                                        //check if they are overlapping
                                        string testOverlappingView = Helpers.CheckVPOverlaps(doc, planViews);
                                        if (testOverlappingView.Contains("xref"))
                                        {
                                            sheetGroup = "Plans Overlapping";
                                            countOverlappingViewports += 1;
                                        }
                                        else
                                        {
                                            sheetGroup = "Plans not Overlapping";
                                            countNonOverlappingViewports += 1;
                                        }
                                        //check if there is a keynote
                                        if (sheetsWithKeynotes.Contains(vs.SheetNumber))
                                        {
                                            sheetGroup += "-Keynote";
                                            keynotesCount += 1;
                                        }
                                    }
                                    else
                                    {
                                        sheetGroup = "No plan views";
                                        countNoPlansViewports += 1;
                                    }

                                    //store the group data in a sheet parameter - hardcoded
                                    Parameter p = vs.LookupParameter(paramExportSheetFilter);

                                    if (p.HasValue && p.AsString() != "")
                                    {
                                        throw new System.ArgumentException();
                                    }
                                    else
                                    {
                                        p.Set(sheetGroup);
                                    }



                                    sb.AppendLine($"{vs.SheetNumber},{sheetGroup}");
                                }//close foreach

                            }//close ProgressForm

                            t.Commit();
                        }
                        File.AppendAllText(outputFile, sb.ToString());

                            TaskDialog myDialog = new TaskDialog("Summary");
                            myDialog.MainIcon = TaskDialogIcon.TaskDialogIconNone;
                            myDialog.MainContent = $"Operation completed.\n{countNoPlansViewports} sheets do not have plan views\n{countNonOverlappingViewports} sheets do not have overlapping views\n{countOverlappingViewports} sheets do have overlapping views\n{keynotesCount} keynotes found";

                            myDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, $"Open Log File {outputFile}", "");

                            TaskDialogResult res = myDialog.Show();

                            if (TaskDialogResult.CommandLink4 == res)
                            {
                                System.Diagnostics.Process process = new System.Diagnostics.Process();
                                process.StartInfo.FileName = outputFile;
                                process.Start();
                            }
                        

                    }//close using form
                }//close try
                catch (System.NullReferenceException)
                {
                    TaskDialog.Show("Error", $"Shared Parameter {paramError}");
                    return Result.Failed;
                }
                catch (System.ArgumentException)
                {
                    TaskDialog.Show("Error", $"The parameter '{paramExportSheetFilter}' is not empty. Operation cancelled.");
                    return Result.Failed;
                }
                catch (System.IO.IOException)
                {
                    TaskDialog.Show("Error", "Please close the log file before exporting.");
                    return Result.Failed;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.GetType().ToString());
                    return Result.Failed;
                }

                return Result.Succeeded;

            }//close form
            catch (Exception ex)
            {
                TaskDialog.Show("Result", ex.Message);
                return Result.Cancelled;
            }
        }//close execute
    }
}
