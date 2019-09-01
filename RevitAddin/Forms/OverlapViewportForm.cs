using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TristanRevitAddin
{
    public partial class OverlapViewportForm : Form
    {

        public List<Autodesk.Revit.DB.ViewSheet> tboxSelectedSheets { get; private set; }
        public List<ViewScheduleOption> cboxSheetDataSource { get; set; }
        private ViewScheduleOption vso = null;

        public OverlapViewportForm()
        {
            InitializeComponent();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {

            tboxSelectedSheets = vso.Views;
        }

        private void OversizedViewportForm_Load(object sender, EventArgs e)
        {
            comboBoxSheetsSchedules.DataSource = cboxSheetDataSource;
            comboBoxSheetsSchedules.DisplayMember = "Name";

        }

        private void comboBoxSheetsSchedules_SelectedIndexChanged(object sender, EventArgs e)
        {
            vso = comboBoxSheetsSchedules.SelectedItem as ViewScheduleOption;
            labelSelectedSheets.Text = $"{vso.ViewSheetCount.ToString()} Sheets selected";
        }
    }
}
