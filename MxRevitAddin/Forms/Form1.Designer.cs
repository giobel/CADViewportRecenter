﻿namespace MxRevitAddin
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.buttonOk = new System.Windows.Forms.Button();
            this.Cancel = new System.Windows.Forms.Button();
            this.textBoxFolder = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.cboxExportSettings = new System.Windows.Forms.ComboBox();
            this.comboBoxSheetsSchedules = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cBoxHideViewportContent = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // buttonOk
            // 
            this.buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOk.Location = new System.Drawing.Point(242, 151);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(81, 25);
            this.buttonOk.TabIndex = 1;
            this.buttonOk.Text = "OK";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.ButtonOk_Click);
            // 
            // Cancel
            // 
            this.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel.Location = new System.Drawing.Point(338, 151);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(84, 25);
            this.Cancel.TabIndex = 2;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            // 
            // textBoxFolder
            // 
            this.textBoxFolder.Location = new System.Drawing.Point(9, 27);
            this.textBoxFolder.Name = "textBoxFolder";
            this.textBoxFolder.Size = new System.Drawing.Size(314, 20);
            this.textBoxFolder.TabIndex = 4;
            this.textBoxFolder.Text = "C:\\Temp\\Metro";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Destination Folder";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 52);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Export Setting";
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(338, 24);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(84, 25);
            this.buttonBrowse.TabIndex = 11;
            this.buttonBrowse.Text = "browse...";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // cboxExportSettings
            // 
            this.cboxExportSettings.FormattingEnabled = true;
            this.cboxExportSettings.Location = new System.Drawing.Point(9, 70);
            this.cboxExportSettings.Name = "cboxExportSettings";
            this.cboxExportSettings.Size = new System.Drawing.Size(224, 21);
            this.cboxExportSettings.TabIndex = 12;
            // 
            // comboBoxSheetsSchedules
            // 
            this.comboBoxSheetsSchedules.FormattingEnabled = true;
            this.comboBoxSheetsSchedules.Location = new System.Drawing.Point(9, 114);
            this.comboBoxSheetsSchedules.Name = "comboBoxSheetsSchedules";
            this.comboBoxSheetsSchedules.Size = new System.Drawing.Size(224, 21);
            this.comboBoxSheetsSchedules.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 96);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "Sheet Schedule";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 207);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(129, 13);
            this.label1.TabIndex = 16;
            this.label1.Text = "version 0.1.2 03/09/2019";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // cBoxHideViewportContent
            // 
            this.cBoxHideViewportContent.AutoSize = true;
            this.cBoxHideViewportContent.Checked = true;
            this.cBoxHideViewportContent.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cBoxHideViewportContent.Location = new System.Drawing.Point(9, 159);
            this.cBoxHideViewportContent.Name = "cBoxHideViewportContent";
            this.cBoxHideViewportContent.Size = new System.Drawing.Size(222, 30);
            this.cBoxHideViewportContent.TabIndex = 17;
            this.cBoxHideViewportContent.Text = "Hide Plan Viewport content \r\n(Warning: linked Keynote will be affected)";
            this.cBoxHideViewportContent.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.cBoxHideViewportContent.UseVisualStyleBackColor = true;
            this.cBoxHideViewportContent.CheckedChanged += new System.EventHandler(this.cBoxHideViewportContent_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(432, 229);
            this.Controls.Add(this.cBoxHideViewportContent);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBoxSheetsSchedules);
            this.Controls.Add(this.cboxExportSettings);
            this.Controls.Add(this.buttonBrowse);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxFolder);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.buttonOk);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Mx CAAD Export";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.TextBox textBoxFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ComboBox cboxExportSettings;
        private System.Windows.Forms.ComboBox comboBoxSheetsSchedules;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox cBoxHideViewportContent;
    }
}