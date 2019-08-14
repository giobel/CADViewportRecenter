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
        public string tBoxSheetNumber { get; set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {
            tBoxSheetNumber = textBox1.Text;
        }
    }
}
