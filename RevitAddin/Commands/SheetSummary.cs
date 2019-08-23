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

namespace RevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class SheetSummary : IExternalCommand
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

                try
                {
                    File.WriteAllText(outputFile,
                      "Sheet Number," +
                      "Group"+
                      Environment.NewLine
                     );

                    List<ViewScheduleOption> viewScheduleOptions = Helpers.GetViewScheduleOptions(doc);

                    using (var form = new OverlapViewportForm())
                    {
                        //set the form sheets
                        form.cboxSheetDataSource = viewScheduleOptions;

                        //use ShowDialog to show the form as a modal dialog box. 
                        form.ShowDialog();

                        //if the user hits cancel just drop out of macro
                        if (form.DialogResult == winForm.DialogResult.Cancel)
                        {
                            return Result.Cancelled;
                        }

                        List<ViewSheet> selectedSheets = form.tboxSelectedSheets;

                        int n = selectedSheets.Count();
                        string s = "{0} of " + n.ToString() + " sheets processed...";
                        string caption = "Find oversized views";

                        int countOverlappingViewports = 0;
                        int countNonOverlappingViewports = 0;
                        int countNoPlansViewports = 0;

                        string sheetGroup = ""; //No plans, Plans not Overlapping, Plans Overlapping

                        using (Transaction t = new Transaction(doc, "Set Sheet Parameter"))
                        {
                            t.Start();

                            using (ProgressForm pf = new ProgressForm(caption, s, n))
                        {
                            foreach (ViewSheet vs in selectedSheets)
                            {
                                if (pf.abortFlag)
                                    break;

                                //select all viewports on the current sheet..not needed TBC
                                ICollection<ElementId> viewports = vs.GetAllViewports();

                                //Select all the views on a sheet
                                ISet<ElementId> placedViewsIds = vs.GetAllPlacedViews();

                                //Filter the views that are plans (Floor,Ceiling,Engineering or Area)
                                List<View> planViews = Helpers.FilterPlanViewport(doc, placedViewsIds);
                                
                                if (planViews.Count > 0) //sheet has plan views, check if they are overlapping
                                {
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
                                }
                                else
                                {
                                    sheetGroup = "No plan views";
                                    countNoPlansViewports += 1;
                                }

                                //store the group data in a sheet parameter - hardcoded
                                    Parameter p = vs.LookupParameter("Mx Export_Sheet Filter");
                                    p.Set(sheetGroup);
                                    
                                
                                sb.AppendLine($"{vs.SheetNumber},{sheetGroup}");
                            }//close foreach
                            
                        }//close ProgressForm

                            t.Commit();
                        }
                        File.AppendAllText(outputFile, sb.ToString());

                        TaskDialog myDialog = new TaskDialog("Summary");
                        myDialog.MainIcon = TaskDialogIcon.TaskDialogIconNone;
                        myDialog.MainContent = $"Operation completed.\n{countNoPlansViewports} sheets do not have plan views\n{countNonOverlappingViewports} sheets do not have overlapping views\n{countOverlappingViewports} sheets do have overlapping views";

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
                    TaskDialog.Show("Error", "Parameter \"Mx Export_Sheet Filter\" not found on Sheet.");
                    return Result.Failed;
                }
                catch (System.IO.IOException)
                {
                    TaskDialog.Show("Error", "Please close the log file before exporting.");
                    return Result.Failed;
                }
                catch(Exception ex)
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
