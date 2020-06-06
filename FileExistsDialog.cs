using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

using InstructorBriefcaseExtractor.Model;

namespace InstructorBriefcaseExtractor
{
    public partial class FileExistsDialog : Form
    {
        private readonly InstructorBriefcaseExtractor.Model.FileInformation myFE;
        public FileExistsDialog(InstructorBriefcaseExtractor.Model.FileInformation FE)
        {
            InitializeComponent();
            myFE = FE;
            try
            {
                txtFilename.Text = myFE.OldFileName;
                txtDirectory.Text = myFE.OldDirectory;

                this.Text = myFE.DialogBoxTitle;
                myFE.CancelALLExport = false;
                myFE.CancelExport = false;
                myFE.OverWriteAll = false;
            }
            catch 
            { }
            
        }

        private void BtnCancelAll_Click(object sender, EventArgs e)
        {
            myFE.CancelALLExport = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            myFE.CancelExport = true;
            this.Close();
        }

        private void BtnOverWriteAll_Click(object sender, EventArgs e)
        {
            myFE.OverWriteAll = true;
            this.Close();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // split the filename apart

            myFE.NewFileName = txtFilename.Text;
            myFE.NewDirectory = txtDirectory.Text;
            this.Close();
        }

        private void FileExistsDialog_Load(object sender, EventArgs e)
        {
            //btnOpenFile.PerformClick();
        }

        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = myFE.OldDirectory;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFilename.Text = Path.GetFileName(openFileDialog1.FileName);
                txtDirectory.Text = Path.GetDirectoryName(openFileDialog1.FileName);
            }
        }
    }
}