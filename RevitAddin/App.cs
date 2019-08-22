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

namespace RevitAddin
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            try
            {
                a.CreateRibbonTab("Metro");

                RibbonPanel tools = a.CreateRibbonPanel("Metro", "Tools");

                AddPushButton(tools, "btnSummary", "Check\nSheets", "", "pack://application:,,,/RevitAddin;component/Images/checkSheet.png", "RevitAddin.SheetSummary", "Report sheets without plan views, with overlapping plans or separated plans.");
                
                AddPushButton(tools, "btnOversized", "Check\nViewport Size", "", "pack://application:,,,/RevitAddin;component/Images/overlap.png", "RevitAddin.FindOversizedViewport", "Report viewports larger than the sheet.");
                //Icon made by Swifticons https://www.flaticon.com is licensed by Creative Commons BY 3.0

                AddPushButton(tools, "btnExportSheets", "Export\nSheets", "", "pack://application:,,,/RevitAddin;component/Images/sheets.png", "RevitAddin.ExportSheets", "Export selected sheets to dwg. View plan elements will be temporarily hidden (atm it only processes 1 plan view per sheet).");

                AddPushButton(tools, "btnTristan", "Export\nXrefs+CSV", "", "pack://application:,,,/RevitAddin;component/Images/xref.png", "RevitAddin.ExportXrefs", "Export selected sheet's plan view to shared coordinates. Creates the csv file for the Autocad script.");
                


                return Result.Succeeded;
            }
            catch
            {
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



 