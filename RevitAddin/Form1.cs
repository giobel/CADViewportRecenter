using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAddin
{
    public partial class Form1 : Form
    {
        public string tBoxDestinationFolder { get; private set; }
        public string tBoxExportSettings { get; private set; }
        public string tBoxSheetNumber { get; private set; }
        public IList<string> cboxDataSource { get; set; }

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cboxExportSettings.DataSource = cboxDataSource;
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {
            tBoxSheetNumber = textBoxSheetNumbers.Text;
            tBoxExportSettings = cboxExportSettings.SelectedValue.ToString();
            tBoxDestinationFolder = textBoxFolder.Text;

        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            textBoxFolder.Text = folderBrowserDialog1.SelectedPath;
        }


    }
}
