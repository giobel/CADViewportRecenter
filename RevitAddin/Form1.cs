using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using winform = System.Windows.Forms;

namespace RevitAddin
{
    public partial class Form1 : winform.Form
    {
        public string tBoxDestinationFolder { get; private set; }
        public string tBoxExportSettings { get; private set; }
        public List<ViewSheet> tboxSelectedSheets { get; private set; }
        public IList<string> cboxExportSettingsDataSource { get; set; }
        public List<ViewScheduleOption> cboxSheetDataSource { get; set; }
        private ViewScheduleOption vso = null;

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cboxExportSettings.DataSource = cboxExportSettingsDataSource;
            comboBoxSheetsSchedules.DataSource = cboxSheetDataSource;
            comboBoxSheetsSchedules.DisplayMember = "Name";
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {
            tboxSelectedSheets = vso.Views;
            tBoxExportSettings = cboxExportSettings.SelectedValue.ToString();
            tBoxDestinationFolder = textBoxFolder.Text;

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
    }
}
