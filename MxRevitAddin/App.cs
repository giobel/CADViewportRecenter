#region Namespaces
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace MxRevitAddin
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            try
            {
                a.CreateRibbonTab("Mx CADD Export");

                RibbonPanel tools = a.CreateRibbonPanel("Mx CADD Export", "Export");

                AddPushButton(tools, "btnSummary", "Check\nSheets", "", "pack://application:,,,/MxRevitAddin;component/Images/checkSheet.png", "MxRevitAddin.CheckSheets", "Report sheets without plan views, with overlapping plans or separated plans.");

                AddPushButton(tools, "btnOversized", "Check\nViewport Size", "", "pack://application:,,,/MxRevitAddin;component/Images/overlap.png", "MxRevitAddin.FindOversizedViewport", "Report viewports larger than the sheet.");
                //Icon made by Swifticons https://www.flaticon.com is licensed by Creative Commons BY 3.0

                AddPushButton(tools, "btnExportSheets", "Export\nSheets", "", "pack://application:,,,/MxRevitAddin;component/Images/sheets.png", "MxRevitAddin.ExportSheets", "Export selected sheets to dwg. View plan elements can be temporarily hidden.");

                AddPushButton(tools, "btnMx", "Export\nXrefs+CSV", "", "pack://application:,,,/MxRevitAddin;component/Images/xref.png", "MxRevitAddin.ExportXrefs", "Export selected sheet's plan view to shared coordinates. Creates the csv file for the Autocad script.");

                AddPushButton(tools, "btnInfo", "Help", "", "pack://application:,,,/MxRevitAddin;component/Images/info.png", "MxRevitAddin.Help", "Read before start exporting.");

                return Result.Succeeded;
            }
            catch(Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        private RibbonPanel GetSetRibbonPanel(UIControlledApplication application, string tabName, string panelName)
        {
            List<RibbonPanel> tabList = new List<RibbonPanel>();

            tabList = application.GetRibbonPanels(tabName);

            RibbonPanel tab = null;

            foreach (RibbonPanel r in tabList)
            {
                if (r.Name.ToUpper() == panelName.ToUpper())
                {
                    tab = r;
                }
            }

            if (tab is null)
                tab = application.CreateRibbonPanel(tabName, panelName);

            return tab;
        }

        private Boolean AddPushButton(RibbonPanel Panel, string ButtonName, string ButtonText, string ImagePath16, string ImagePath32, string dllClass, string Tooltip)
        {

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            try
            {
                PushButtonData m_pbData = new PushButtonData(ButtonName, ButtonText, thisAssemblyPath, dllClass);

                if (ImagePath16 != "")
                {
                    try
                    {
                        m_pbData.Image = new BitmapImage(new Uri(ImagePath16));
                    }
                    catch
                    {
                        //Could not find the image
                    }
                }
                if (ImagePath32 != "")
                {
                    try
                    {
                        m_pbData.LargeImage = new BitmapImage(new Uri(ImagePath32));
                    }
                    catch
                    {
                        //Could not find the image
                    }
                }

                m_pbData.ToolTip = Tooltip;


                PushButton m_pb = Panel.AddItem(m_pbData) as PushButton;

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}



