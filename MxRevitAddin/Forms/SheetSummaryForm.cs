using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MxRevitAddin
{
    public partial class SheetSummaryForm : Form
    {
        public List<Autodesk.Revit.DB.ViewSchedule> cboxSheetDataSource { get; set; }
        public Autodesk.Revit.DB.ViewSchedule selectedViewSchedule = null;

        public SheetSummaryForm()
        {
            InitializeComponent();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {

            selectedViewSchedule = comboBoxSheetsSchedules.SelectedItem as Autodesk.Revit.DB.ViewSchedule;
        }

        private void OversizedViewportForm_Load(object sender, EventArgs e)
        {
            comboBoxSheetsSchedules.DataSource = cboxSheetDataSource;
            comboBoxSheetsSchedules.DisplayMember = "Name";

        }

        //private void comboBoxSheetsSchedules_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    var vso = comboBoxSheetsSchedules.SelectedItem as Autodesk.Revit.DB.ViewSchedule;
        //    labelSelectedSheets.Text = $"{viewSheetCount} Sheets selected";
        //}
    }
}
