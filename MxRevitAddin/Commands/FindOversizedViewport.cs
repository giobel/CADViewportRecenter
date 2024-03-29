﻿using Autodesk.Revit.ApplicationServices;
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
    public class FindOversizedViewport : IExternalCommand
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

                string outputFile = @"C:\Temp\reportOversizedVP.csv";

                StringBuilder sb = new StringBuilder();

                try
                {
                    File.WriteAllText(outputFile,
                      "Sheet Number," +
                      "Viewport Name," +
                      Environment.NewLine
                     );
                }
                catch
                {
                    TaskDialog.Show("Error", "Operation cancelled. Please close the log file.");
                    return Result.Failed;
                }

                //List<ViewScheduleOption> viewScheduleOptions = Helpers.GetViewScheduleOptions(doc);

                List<ViewSchedule> viewSchedule = Helpers.SheetList(doc);



                using (var form = new OversizedViewportForm())
                {
                    //set the form sheets
                    form.cboxSheetDataSource = viewSchedule;

                    //se the form title
                    form.Text = "Check Viewport Oversized";

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

                    int maxWidth = form.maxWidth;
                    int maxHeight = form.maxHeight;
                    int centerX = form.centerX;
                    int centerY = form.centerY;

                    int n = selectedSheets.Count();
                    string s = "{0} of " + n.ToString() + " sheets processed...";
                    string caption = "Find oversized views";

                    int countOversizedViewports = 0;

                    using (ProgressForm pf = new ProgressForm(caption, s, n))
                    {

                        foreach (ViewSheet vs in selectedSheets)
                        {
                            if (pf.abortFlag)
                                break;

                            ICollection<ElementId> viewports = vs.GetAllViewports();

                            ISet<ElementId> placedViewsIds = vs.GetAllPlacedViews();

                            List<View> planViews = Helpers.FilterPlanViewport(doc, placedViewsIds);
                            
                            foreach (ElementId vpId in viewports)
                            {
                                try
                                {
                                    Viewport vp = doc.GetElement(vpId) as Viewport;
                                    View planView = doc.GetElement(vp.ViewId) as View;


                                    if (Helpers.ViewportIntersectsTitleblock(vp, centerX, centerY, maxWidth, maxHeight))
                                    {
                                        sb.AppendLine($"{vs.SheetNumber},{planView.Name}");
                                        countOversizedViewports += 1;
                                    }

                                    //if (planView.ViewType == ViewType.FloorPlan || planView.ViewType == ViewType.EngineeringPlan || planView.ViewType == ViewType.CeilingPlan)
                                    //{
                                        //XYZ maxPt = vp.GetBoxOutline().MaximumPoint; too slow
                                        //XYZ minPt = vp.GetBoxOutline().MinimumPoint;
                                        //BoundingBoxXYZ bbox = vp.get_BoundingBox(vs);

                                        //XYZ maxPt = bbox.Max;
                                        //XYZ minPt = bbox.Min;

                                        //int width = Convert.ToInt32((maxPt.X - minPt.X) * 304.8);
                                        //int height = Convert.ToInt32((maxPt.Y - minPt.Y) * 304.8);

                                        //int checkWidth = maxWidth;
                                        //int checkHeight = maxHeight;

                                        //if (width > checkWidth || height > checkHeight)
                                        //{
                                        //    sb.AppendLine($"{vs.SheetNumber},{vp.Name},{width},{height}");
                                        //    countOversizedViewports += 1;
                                        //}       

                                    //    if (Helpers.ViewportIntersectsTitleblock(vp))
                                    //    {
                                    //        sb.AppendLine($"{vs.SheetNumber},{vp.Name}");
                                    //        countOversizedViewports += 1;
                                    //    }
                                    //}
                                }
                                catch
                                {
                                    sb.AppendLine($"{vs.SheetNumber}, ERROR");
                                }

                            }
                            
                            pf.Increment();
                        }//close foreach
                    }//close progress form

                    File.AppendAllText(outputFile, sb.ToString());

                    TaskDialog myDialog = new TaskDialog("Summary");
                    myDialog.MainIcon = TaskDialogIcon.TaskDialogIconNone;
                    myDialog.MainContent = $"Operation completed. {countOversizedViewports} viewports are oversized.";

                    myDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, $"Open Log File {outputFile}", "");

                    TaskDialogResult res = myDialog.Show();

                    if (TaskDialogResult.CommandLink4 == res)
                    {
                        System.Diagnostics.Process process = new System.Diagnostics.Process();
                        process.StartInfo.FileName = outputFile;
                        process.Start();
                    }
                    return Result.Succeeded;
                }//close form
            }//close try
            catch (Exception ex)
            {
                TaskDialog.Show("Result", ex.Message);
                return Result.Cancelled;
            }
        }//close execute
    }
}
