using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MxRevitAddin
{
    public partial class OversizedViewportForm : Form
    {
        public int maxWidth { get; private set; }
        public int maxHeight { get; private set; }
        public int centerX { get; private set; }
        public int centerY { get; private set; }
        public List<Autodesk.Revit.DB.ViewSchedule> cboxSheetDataSource { get; set; }
        public Autodesk.Revit.DB.ViewSchedule selectedViewSchedule = null;

        public OversizedViewportForm()
        {
            InitializeComponent();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            maxWidth = Convert.ToInt16(textBoxWidth.Text);
            maxHeight = Convert.ToInt16(textBoxHeight.Text);
            centerX = Convert.ToInt16(textBoxCenterX.Text);
            centerY = Convert.ToInt16(textBoxCenterY.Text);
            selectedViewSchedule = comboBoxSheetsSchedules.SelectedItem as Autodesk.Revit.DB.ViewSchedule;
        }

        private void OversizedViewportForm_Load(object sender, EventArgs e)
        {
            comboBoxSheetsSchedules.DataSource = cboxSheetDataSource;
            comboBoxSheetsSchedules.DisplayMember = "Name";

        }

        //private void comboBoxSheetsSchedules_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    vso = comboBoxSheetsSchedules.SelectedItem as ViewScheduleOption;
        //    labelSelectedSheets.Text = $"{vso.ViewSheetCount.ToString()} Sheets selected";
        //}
    }
}
