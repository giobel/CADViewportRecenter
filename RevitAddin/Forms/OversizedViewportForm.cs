using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TristanRevitAddin
{
    public partial class OversizedViewportForm : Form
    {
        public int maxWidth { get; private set; }
        public int maxHeight { get; private set; }
        public List<Autodesk.Revit.DB.ViewSheet> tboxSelectedSheets { get; private set; }
        public List<ViewScheduleOption> cboxSheetDataSource { get; set; }
        private ViewScheduleOption vso = null;

        public OversizedViewportForm()
        {
            InitializeComponent();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            maxWidth = Convert.ToInt16(textBoxWidth.Text);
            maxHeight = Convert.ToInt16(textBoxHeight.Text);
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
