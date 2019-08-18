﻿using System;
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
    public partial class ProgressForm : Form
    {
        public bool abortFlag { get; private set; }
        string _format;
        public ProgressForm(string caption, string format, int max)
        {
            _format = format;
            InitializeComponent();
            Text = caption;
            label1.Text = (null == format) ? caption : string.Format(format, 0);
            progressBar1.Minimum = 0;
            progressBar1.Maximum = max;
            progressBar1.Value = 0;
            Show();
            Application.DoEvents();
        }

        public void Increment()
        {
            ++progressBar1.Value;

            if (null != _format)
            {
                label1.Text = string.Format(_format, progressBar1.Value);
            }
            Application.DoEvents();
        }

        private void ButtonAbort_Click(object sender, EventArgs e)
        {
            label1.Text = "Aborting...";
            abortFlag = true;
        }
    }
}
