using System;
using System.Collections.Generic;
using winform = System.Windows.Forms;

namespace TristanRevitAddin
{
    public partial class Form1 : winform.Form
    {
        public string TBoxDestinationFolder { get; private set; }
        public string TBoxExportSettings { get; private set; }
        public List<Autodesk.Revit.DB.ViewSheet> TboxSelectedSheets { get; private set; }
        public IList<string> CboxExportSettingsDataSource { get; set; }
        public List<ViewScheduleOption> CboxSheetDataSource { get; set; }
        private ViewScheduleOption vso = null;
        public bool HideViewportContent { get; private set; }

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cboxExportSettings.DataSource = CboxExportSettingsDataSource;
            comboBoxSheetsSchedules.DataSource = CboxSheetDataSource;
            comboBoxSheetsSchedules.DisplayMember = "Name";
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {
            TboxSelectedSheets = vso.Views;
            TBoxExportSettings = cboxExportSettings.SelectedValue.ToString();
            TBoxDestinationFolder = textBoxFolder.Text;
            HideViewportContent = cBoxHideViewportContent.Checked;
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            winform.DialogResult result = folderBrowserDialog1.ShowDialog();
            textBoxFolder.Text = folderBrowserDialog1.SelectedPath;
        }

        private void comboBoxSheetsSchedules_SelectedIndexChanged(object sender, EventArgs e)
        {
            vso = comboBoxSheetsSchedules.SelectedItem as ViewScheduleOption;
            labelSelectedSheets.Text = $"{vso.ViewSheetCount.ToString()} Sheets selected";
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Autodesk.Revit.UI.TaskDialog.Show("Hey", "What where you trying to do??");
        }

        private void cBoxHideViewportContent_CheckedChanged(object sender, EventArgs e)
        {
            
        }
    }
}
