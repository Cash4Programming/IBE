namespace InstructorBriefcaseExtractor
{
    partial class Extractor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Extractor));
            this.BtnMoreLess = new System.Windows.Forms.Button();
            this.BtnBack = new System.Windows.Forms.Button();
            this.BtnNext = new System.Windows.Forms.Button();
            this.RtbHTML = new System.Windows.Forms.RichTextBox();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.PanelCourseRetrieved = new System.Windows.Forms.Panel();
            this.GB3Wamap = new System.Windows.Forms.GroupBox();
            this.CbkWamapExport = new System.Windows.Forms.CheckBox();
            this.GB3Clicker = new System.Windows.Forms.GroupBox();
            this.CbkClickerExport = new System.Windows.Forms.CheckBox();
            this.ChkOverWrite = new System.Windows.Forms.CheckBox();
            this.GB3WebAssign = new System.Windows.Forms.GroupBox();
            this.CbkWebAssignExport = new System.Windows.Forms.CheckBox();
            this.GB3Outlook = new System.Windows.Forms.GroupBox();
            this.CbkOutlookExport = new System.Windows.Forms.CheckBox();
            this.GB3Excel = new System.Windows.Forms.GroupBox();
            this.CbkExcelLabExport = new System.Windows.Forms.CheckBox();
            this.CbkExcelRollExport = new System.Windows.Forms.CheckBox();
            this.CbkExcelClassExport = new System.Windows.Forms.CheckBox();
            this.GB3MTG = new System.Windows.Forms.GroupBox();
            this.CbkMTGExport = new System.Windows.Forms.CheckBox();
            this.ClbCourses = new System.Windows.Forms.CheckedListBox();
            this.LblCourses = new System.Windows.Forms.Label();
            this.PanelQuarterFound = new System.Windows.Forms.Panel();
            this.LblYRQ = new System.Windows.Forms.Label();
            this.CmbQuarterName = new System.Windows.Forms.ComboBox();
            this.LblAddress2 = new System.Windows.Forms.Label();
            this.TxtAddress = new System.Windows.Forms.TextBox();
            this.TxtPIN = new System.Windows.Forms.TextBox();
            this.TxtUserName = new System.Windows.Forms.TextBox();
            this.LblPassword = new System.Windows.Forms.Label();
            this.LblEmployeeID = new System.Windows.Forms.Label();
            this.PanelOutput = new System.Windows.Forms.Panel();
            this.LblOutputSummary = new System.Windows.Forms.Label();
            this.RtbOutput = new System.Windows.Forms.RichTextBox();
            this.LblVersion = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ExitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ConfigurationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.CbkExcelRollWaitListExport = new System.Windows.Forms.CheckBox();
            this.CbkExcelClassWaitListExport = new System.Windows.Forms.CheckBox();
            this.PanelCourseRetrieved.SuspendLayout();
            this.GB3Wamap.SuspendLayout();
            this.GB3Clicker.SuspendLayout();
            this.GB3WebAssign.SuspendLayout();
            this.GB3Outlook.SuspendLayout();
            this.GB3Excel.SuspendLayout();
            this.GB3MTG.SuspendLayout();
            this.PanelQuarterFound.SuspendLayout();
            this.PanelOutput.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnMoreLess
            // 
            this.BtnMoreLess.Location = new System.Drawing.Point(406, 385);
            this.BtnMoreLess.Name = "BtnMoreLess";
            this.BtnMoreLess.Size = new System.Drawing.Size(80, 24);
            this.BtnMoreLess.TabIndex = 25;
            this.BtnMoreLess.Text = "More >>";
            this.BtnMoreLess.Click += new System.EventHandler(this.BtnMoreLess_Click);
            // 
            // BtnBack
            // 
            this.BtnBack.Location = new System.Drawing.Point(214, 353);
            this.BtnBack.Name = "BtnBack";
            this.BtnBack.Size = new System.Drawing.Size(80, 24);
            this.BtnBack.TabIndex = 24;
            this.BtnBack.Text = "Back";
            this.BtnBack.Click += new System.EventHandler(this.BtnBack_Click);
            // 
            // BtnNext
            // 
            this.BtnNext.Location = new System.Drawing.Point(310, 353);
            this.BtnNext.Name = "BtnNext";
            this.BtnNext.Size = new System.Drawing.Size(80, 24);
            this.BtnNext.TabIndex = 23;
            this.BtnNext.Text = "Next";
            this.BtnNext.Click += new System.EventHandler(this.BtnNext_Click);
            // 
            // RtbHTML
            // 
            this.RtbHTML.Location = new System.Drawing.Point(12, 426);
            this.RtbHTML.Name = "RtbHTML";
            this.RtbHTML.Size = new System.Drawing.Size(480, 211);
            this.RtbHTML.TabIndex = 22;
            this.RtbHTML.Text = "";
            // 
            // BtnCancel
            // 
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(406, 353);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(80, 24);
            this.BtnCancel.TabIndex = 21;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // PanelCourseRetrieved
            // 
            this.PanelCourseRetrieved.Controls.Add(this.GB3Wamap);
            this.PanelCourseRetrieved.Controls.Add(this.GB3Clicker);
            this.PanelCourseRetrieved.Controls.Add(this.ChkOverWrite);
            this.PanelCourseRetrieved.Controls.Add(this.GB3WebAssign);
            this.PanelCourseRetrieved.Controls.Add(this.GB3Outlook);
            this.PanelCourseRetrieved.Controls.Add(this.GB3Excel);
            this.PanelCourseRetrieved.Controls.Add(this.GB3MTG);
            this.PanelCourseRetrieved.Controls.Add(this.ClbCourses);
            this.PanelCourseRetrieved.Controls.Add(this.LblCourses);
            this.PanelCourseRetrieved.Location = new System.Drawing.Point(10, 25);
            this.PanelCourseRetrieved.Name = "PanelCourseRetrieved";
            this.PanelCourseRetrieved.Size = new System.Drawing.Size(522, 322);
            this.PanelCourseRetrieved.TabIndex = 27;
            // 
            // GB3Wamap
            // 
            this.GB3Wamap.Controls.Add(this.CbkWamapExport);
            this.GB3Wamap.Location = new System.Drawing.Point(212, 55);
            this.GB3Wamap.Name = "GB3Wamap";
            this.GB3Wamap.Size = new System.Drawing.Size(150, 50);
            this.GB3Wamap.TabIndex = 9;
            this.GB3Wamap.TabStop = false;
            this.GB3Wamap.Text = "WAMAP";
            // 
            // CbkWamapExport
            // 
            this.CbkWamapExport.Location = new System.Drawing.Point(8, 16);
            this.CbkWamapExport.Name = "CbkWamapExport";
            this.CbkWamapExport.Size = new System.Drawing.Size(91, 28);
            this.CbkWamapExport.TabIndex = 1;
            this.CbkWamapExport.Text = "Export";
            this.CbkWamapExport.CheckedChanged += new System.EventHandler(this.CbkWamapExport_CheckedChanged);
            // 
            // GB3Clicker
            // 
            this.GB3Clicker.Controls.Add(this.CbkClickerExport);
            this.GB3Clicker.Location = new System.Drawing.Point(212, 7);
            this.GB3Clicker.Name = "GB3Clicker";
            this.GB3Clicker.Size = new System.Drawing.Size(150, 52);
            this.GB3Clicker.TabIndex = 10;
            this.GB3Clicker.TabStop = false;
            this.GB3Clicker.Text = "Clicker";
            // 
            // CbkClickerExport
            // 
            this.CbkClickerExport.Location = new System.Drawing.Point(8, 16);
            this.CbkClickerExport.Name = "CbkClickerExport";
            this.CbkClickerExport.Size = new System.Drawing.Size(91, 30);
            this.CbkClickerExport.TabIndex = 1;
            this.CbkClickerExport.Text = "Export";
            this.CbkClickerExport.CheckedChanged += new System.EventHandler(this.CbkClickerExport_CheckedChanged);
            // 
            // ChkOverWrite
            // 
            this.ChkOverWrite.AutoSize = true;
            this.ChkOverWrite.Location = new System.Drawing.Point(212, 289);
            this.ChkOverWrite.Name = "ChkOverWrite";
            this.ChkOverWrite.Size = new System.Drawing.Size(255, 21);
            this.ChkOverWrite.TabIndex = 9;
            this.ChkOverWrite.Text = "Replace existing files without asking";
            this.ChkOverWrite.UseVisualStyleBackColor = true;
            // 
            // GB3WebAssign
            // 
            this.GB3WebAssign.Controls.Add(this.CbkWebAssignExport);
            this.GB3WebAssign.Location = new System.Drawing.Point(215, 113);
            this.GB3WebAssign.Name = "GB3WebAssign";
            this.GB3WebAssign.Size = new System.Drawing.Size(127, 55);
            this.GB3WebAssign.TabIndex = 8;
            this.GB3WebAssign.TabStop = false;
            this.GB3WebAssign.Text = "WebAssign";
            // 
            // CbkWebAssignExport
            // 
            this.CbkWebAssignExport.Location = new System.Drawing.Point(8, 16);
            this.CbkWebAssignExport.Name = "CbkWebAssignExport";
            this.CbkWebAssignExport.Size = new System.Drawing.Size(90, 29);
            this.CbkWebAssignExport.TabIndex = 1;
            this.CbkWebAssignExport.Text = "Export";
            this.CbkWebAssignExport.CheckedChanged += new System.EventHandler(this.CbkWebAssignExport_CheckedChanged);
            // 
            // GB3Outlook
            // 
            this.GB3Outlook.Controls.Add(this.CbkOutlookExport);
            this.GB3Outlook.Location = new System.Drawing.Point(368, 9);
            this.GB3Outlook.Name = "GB3Outlook";
            this.GB3Outlook.Size = new System.Drawing.Size(127, 56);
            this.GB3Outlook.TabIndex = 6;
            this.GB3Outlook.TabStop = false;
            this.GB3Outlook.Text = "Outlook";
            // 
            // CbkOutlookExport
            // 
            this.CbkOutlookExport.Location = new System.Drawing.Point(8, 18);
            this.CbkOutlookExport.Name = "CbkOutlookExport";
            this.CbkOutlookExport.Size = new System.Drawing.Size(91, 32);
            this.CbkOutlookExport.TabIndex = 1;
            this.CbkOutlookExport.Text = "Export";
            this.CbkOutlookExport.CheckedChanged += new System.EventHandler(this.CbkOutlookExport_CheckedChanged);
            // 
            // GB3Excel
            // 
            this.GB3Excel.Controls.Add(this.CbkExcelClassWaitListExport);
            this.GB3Excel.Controls.Add(this.CbkExcelRollWaitListExport);
            this.GB3Excel.Controls.Add(this.CbkExcelLabExport);
            this.GB3Excel.Controls.Add(this.CbkExcelRollExport);
            this.GB3Excel.Controls.Add(this.CbkExcelClassExport);
            this.GB3Excel.Location = new System.Drawing.Point(223, 167);
            this.GB3Excel.Name = "GB3Excel";
            this.GB3Excel.Size = new System.Drawing.Size(274, 116);
            this.GB3Excel.TabIndex = 5;
            this.GB3Excel.TabStop = false;
            this.GB3Excel.Text = "Excel";
            // 
            // CbkExcelLabExport
            // 
            this.CbkExcelLabExport.Location = new System.Drawing.Point(8, 79);
            this.CbkExcelLabExport.Name = "CbkExcelLabExport";
            this.CbkExcelLabExport.Size = new System.Drawing.Size(130, 24);
            this.CbkExcelLabExport.TabIndex = 3;
            this.CbkExcelLabExport.Text = "Export Lab";
            // 
            // CbkExcelRollExport
            // 
            this.CbkExcelRollExport.Location = new System.Drawing.Point(8, 48);
            this.CbkExcelRollExport.Name = "CbkExcelRollExport";
            this.CbkExcelRollExport.Size = new System.Drawing.Size(128, 23);
            this.CbkExcelRollExport.TabIndex = 2;
            this.CbkExcelRollExport.Text = "Export Roll";
            // 
            // CbkExcelClassExport
            // 
            this.CbkExcelClassExport.Location = new System.Drawing.Point(8, 16);
            this.CbkExcelClassExport.Name = "CbkExcelClassExport";
            this.CbkExcelClassExport.Size = new System.Drawing.Size(128, 26);
            this.CbkExcelClassExport.TabIndex = 1;
            this.CbkExcelClassExport.Text = "Export Class";
            this.CbkExcelClassExport.CheckedChanged += new System.EventHandler(this.CbkExcelClassExport_CheckedChanged);
            // 
            // GB3MTG
            // 
            this.GB3MTG.Controls.Add(this.CbkMTGExport);
            this.GB3MTG.Location = new System.Drawing.Point(370, 63);
            this.GB3MTG.Name = "GB3MTG";
            this.GB3MTG.Size = new System.Drawing.Size(127, 50);
            this.GB3MTG.TabIndex = 4;
            this.GB3MTG.TabStop = false;
            this.GB3MTG.Text = "MTG";
            // 
            // CbkMTGExport
            // 
            this.CbkMTGExport.Location = new System.Drawing.Point(8, 16);
            this.CbkMTGExport.Name = "CbkMTGExport";
            this.CbkMTGExport.Size = new System.Drawing.Size(89, 28);
            this.CbkMTGExport.TabIndex = 1;
            this.CbkMTGExport.Text = "Export";
            this.CbkMTGExport.CheckedChanged += new System.EventHandler(this.CbkMTGExport_CheckedChanged);
            // 
            // ClbCourses
            // 
            this.ClbCourses.Location = new System.Drawing.Point(8, 32);
            this.ClbCourses.Name = "ClbCourses";
            this.ClbCourses.Size = new System.Drawing.Size(196, 276);
            this.ClbCourses.TabIndex = 3;
            this.ClbCourses.SelectedIndexChanged += new System.EventHandler(this.ClbCourses_SelectedIndexChanged);
            // 
            // LblCourses
            // 
            this.LblCourses.Location = new System.Drawing.Point(5, 7);
            this.LblCourses.Name = "LblCourses";
            this.LblCourses.Size = new System.Drawing.Size(183, 16);
            this.LblCourses.TabIndex = 0;
            this.LblCourses.Text = "Select WINTER 2008 Course(s) to Export";
            // 
            // PanelQuarterFound
            // 
            this.PanelQuarterFound.Controls.Add(this.LblYRQ);
            this.PanelQuarterFound.Controls.Add(this.CmbQuarterName);
            this.PanelQuarterFound.Controls.Add(this.LblAddress2);
            this.PanelQuarterFound.Controls.Add(this.TxtAddress);
            this.PanelQuarterFound.Controls.Add(this.TxtPIN);
            this.PanelQuarterFound.Controls.Add(this.TxtUserName);
            this.PanelQuarterFound.Controls.Add(this.LblPassword);
            this.PanelQuarterFound.Controls.Add(this.LblEmployeeID);
            this.PanelQuarterFound.Location = new System.Drawing.Point(571, 32);
            this.PanelQuarterFound.Name = "PanelQuarterFound";
            this.PanelQuarterFound.Size = new System.Drawing.Size(513, 293);
            this.PanelQuarterFound.TabIndex = 28;
            // 
            // LblYRQ
            // 
            this.LblYRQ.Location = new System.Drawing.Point(121, 140);
            this.LblYRQ.Name = "LblYRQ";
            this.LblYRQ.Size = new System.Drawing.Size(100, 16);
            this.LblYRQ.TabIndex = 19;
            this.LblYRQ.Text = "Quarter";
            // 
            // CmbQuarterName
            // 
            this.CmbQuarterName.Location = new System.Drawing.Point(245, 140);
            this.CmbQuarterName.MaxDropDownItems = 3;
            this.CmbQuarterName.Name = "CmbQuarterName";
            this.CmbQuarterName.Size = new System.Drawing.Size(168, 24);
            this.CmbQuarterName.TabIndex = 18;
            // 
            // LblAddress2
            // 
            this.LblAddress2.Location = new System.Drawing.Point(19, 22);
            this.LblAddress2.Name = "LblAddress2";
            this.LblAddress2.Size = new System.Drawing.Size(75, 16);
            this.LblAddress2.TabIndex = 17;
            this.LblAddress2.Text = "Web Site";
            // 
            // TxtAddress
            // 
            this.TxtAddress.Enabled = false;
            this.TxtAddress.Location = new System.Drawing.Point(106, 22);
            this.TxtAddress.Name = "TxtAddress";
            this.TxtAddress.Size = new System.Drawing.Size(388, 22);
            this.TxtAddress.TabIndex = 16;
            // 
            // TxtPIN
            // 
            this.TxtPIN.Location = new System.Drawing.Point(245, 116);
            this.TxtPIN.Name = "TxtPIN";
            this.TxtPIN.PasswordChar = '*';
            this.TxtPIN.Size = new System.Drawing.Size(168, 22);
            this.TxtPIN.TabIndex = 15;
            // 
            // TxtUserName
            // 
            this.TxtUserName.Location = new System.Drawing.Point(245, 92);
            this.TxtUserName.Name = "TxtUserName";
            this.TxtUserName.Size = new System.Drawing.Size(168, 22);
            this.TxtUserName.TabIndex = 12;
            // 
            // LblPassword
            // 
            this.LblPassword.Location = new System.Drawing.Point(121, 116);
            this.LblPassword.Name = "LblPassword";
            this.LblPassword.Size = new System.Drawing.Size(100, 16);
            this.LblPassword.TabIndex = 14;
            this.LblPassword.Text = "Employee PIN";
            // 
            // LblEmployeeID
            // 
            this.LblEmployeeID.Location = new System.Drawing.Point(121, 92);
            this.LblEmployeeID.Name = "LblEmployeeID";
            this.LblEmployeeID.Size = new System.Drawing.Size(100, 16);
            this.LblEmployeeID.TabIndex = 13;
            this.LblEmployeeID.Text = "Employee ID";
            // 
            // PanelOutput
            // 
            this.PanelOutput.Controls.Add(this.LblOutputSummary);
            this.PanelOutput.Controls.Add(this.RtbOutput);
            this.PanelOutput.Location = new System.Drawing.Point(571, 361);
            this.PanelOutput.Name = "PanelOutput";
            this.PanelOutput.Size = new System.Drawing.Size(513, 293);
            this.PanelOutput.TabIndex = 29;
            // 
            // LblOutputSummary
            // 
            this.LblOutputSummary.AutoSize = true;
            this.LblOutputSummary.Location = new System.Drawing.Point(9, 4);
            this.LblOutputSummary.Name = "LblOutputSummary";
            this.LblOutputSummary.Size = new System.Drawing.Size(114, 17);
            this.LblOutputSummary.TabIndex = 1;
            this.LblOutputSummary.Text = "Output Summary";
            // 
            // RtbOutput
            // 
            this.RtbOutput.Location = new System.Drawing.Point(7, 20);
            this.RtbOutput.Name = "RtbOutput";
            this.RtbOutput.Size = new System.Drawing.Size(487, 217);
            this.RtbOutput.TabIndex = 0;
            this.RtbOutput.Text = "";
            // 
            // LblVersion
            // 
            this.LblVersion.AutoSize = true;
            this.LblVersion.Location = new System.Drawing.Point(7, 391);
            this.LblVersion.Name = "LblVersion";
            this.LblVersion.Size = new System.Drawing.Size(56, 17);
            this.LblVersion.TabIndex = 30;
            this.LblVersion.Text = "Version";
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.ConfigurationToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(982, 28);
            this.menuStrip1.TabIndex = 31;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ExitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(44, 24);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // ExitToolStripMenuItem
            // 
            this.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem";
            this.ExitToolStripMenuItem.Size = new System.Drawing.Size(108, 26);
            this.ExitToolStripMenuItem.Text = "Exit";
            this.ExitToolStripMenuItem.Click += new System.EventHandler(this.ExitToolStripMenuItem_Click);
            // 
            // ConfigurationToolStripMenuItem
            // 
            this.ConfigurationToolStripMenuItem.Name = "ConfigurationToolStripMenuItem";
            this.ConfigurationToolStripMenuItem.Size = new System.Drawing.Size(112, 24);
            this.ConfigurationToolStripMenuItem.Text = "Configuration";
            this.ConfigurationToolStripMenuItem.Click += new System.EventHandler(this.ConfigurationToolStripMenuItem_Click);
            // 
            // CbkExcelRollWaitListExport
            // 
            this.CbkExcelRollWaitListExport.Location = new System.Drawing.Point(129, 49);
            this.CbkExcelRollWaitListExport.Name = "CbkExcelRollWaitListExport";
            this.CbkExcelRollWaitListExport.Size = new System.Drawing.Size(130, 24);
            this.CbkExcelRollWaitListExport.TabIndex = 5;
            this.CbkExcelRollWaitListExport.Text = "Include Waitlist";
            // 
            // CbkExcelClassWaitListExport
            // 
            this.CbkExcelClassWaitListExport.Location = new System.Drawing.Point(129, 17);
            this.CbkExcelClassWaitListExport.Name = "CbkExcelClassWaitListExport";
            this.CbkExcelClassWaitListExport.Size = new System.Drawing.Size(130, 24);
            this.CbkExcelClassWaitListExport.TabIndex = 6;
            this.CbkExcelClassWaitListExport.Text = "Include Waitlist";
            // 
            // Extractor
            // 
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(982, 645);
            this.Controls.Add(this.LblVersion);
            this.Controls.Add(this.PanelOutput);
            this.Controls.Add(this.PanelQuarterFound);
            this.Controls.Add(this.PanelCourseRetrieved);
            this.Controls.Add(this.BtnMoreLess);
            this.Controls.Add(this.BtnBack);
            this.Controls.Add(this.BtnNext);
            this.Controls.Add(this.RtbHTML);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Extractor";
            this.Text = "Instructor Briefcase Extractor";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Extractor_FormClosed);
            this.PanelCourseRetrieved.ResumeLayout(false);
            this.PanelCourseRetrieved.PerformLayout();
            this.GB3Wamap.ResumeLayout(false);
            this.GB3Clicker.ResumeLayout(false);
            this.GB3WebAssign.ResumeLayout(false);
            this.GB3Outlook.ResumeLayout(false);
            this.GB3Excel.ResumeLayout(false);
            this.GB3MTG.ResumeLayout(false);
            this.PanelQuarterFound.ResumeLayout(false);
            this.PanelQuarterFound.PerformLayout();
            this.PanelOutput.ResumeLayout(false);
            this.PanelOutput.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Button BtnMoreLess;
        internal System.Windows.Forms.Button BtnBack;
        internal System.Windows.Forms.Button BtnNext;
        private System.Windows.Forms.RichTextBox RtbHTML;
        internal System.Windows.Forms.Button BtnCancel;
        internal System.Windows.Forms.Panel PanelCourseRetrieved;
        internal System.Windows.Forms.GroupBox GB3WebAssign;
        internal System.Windows.Forms.CheckBox CbkWebAssignExport;
        internal System.Windows.Forms.GroupBox GB3Outlook;
        internal System.Windows.Forms.CheckBox CbkOutlookExport;
        internal System.Windows.Forms.GroupBox GB3Excel;
        internal System.Windows.Forms.CheckBox CbkExcelClassExport;
        internal System.Windows.Forms.GroupBox GB3MTG;
        internal System.Windows.Forms.CheckBox CbkMTGExport;
        internal System.Windows.Forms.CheckedListBox ClbCourses;
        internal System.Windows.Forms.Label LblCourses;
        internal System.Windows.Forms.Panel PanelQuarterFound;
        internal System.Windows.Forms.Label LblYRQ;
        internal System.Windows.Forms.ComboBox CmbQuarterName;
        internal System.Windows.Forms.Label LblAddress2;
        internal System.Windows.Forms.TextBox TxtAddress;
        internal System.Windows.Forms.TextBox TxtPIN;
        internal System.Windows.Forms.TextBox TxtUserName;
        internal System.Windows.Forms.Label LblPassword;
        internal System.Windows.Forms.Label LblEmployeeID;
        private System.Windows.Forms.Panel PanelOutput;
        private System.Windows.Forms.CheckBox ChkOverWrite;
        private System.Windows.Forms.Label LblOutputSummary;
        private System.Windows.Forms.RichTextBox RtbOutput;
        private System.Windows.Forms.Label LblVersion;
        internal System.Windows.Forms.GroupBox GB3Clicker;
        internal System.Windows.Forms.CheckBox CbkClickerExport;
        internal System.Windows.Forms.GroupBox GB3Wamap;
        internal System.Windows.Forms.CheckBox CbkWamapExport;
        internal System.Windows.Forms.CheckBox CbkExcelRollExport;
        internal System.Windows.Forms.CheckBox CbkExcelLabExport;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ExitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ConfigurationToolStripMenuItem;
        internal System.Windows.Forms.CheckBox CbkExcelClassWaitListExport;
        internal System.Windows.Forms.CheckBox CbkExcelRollWaitListExport;
    }
}

