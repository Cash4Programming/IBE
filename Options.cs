using System;
using System.IO;
using System.Windows.Forms;

using InstructorBriefcaseExtractor.BLL;
using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.DAL;

namespace InstructorBriefcaseExtractor
{
    public partial class Options : Form
    {
        #region Properties

        public UserSettings UserSettings { get; set; }
        public ClickerSettings ClickerSettings { get; set; }
        public MTGSettings MTGSettings { get; set; }
        public OutlookSettings OutlookSettings { get; set; }
        public WamapSettings WamapSettings { get; set; }
        public WebAssignSettings WebAssignSettings { get; set; }
        public ExcelRollSettings ExcelRollSettings { get; set; }
        public ExcelClassSettings ExcelClassSettings { get; set; }
        public Colleges Colleges { get; set; }
        public Quarter2MonthSettings Quarter2Month { get; set; }

        #endregion

        private bool CollegeOnly = false;
        private GenericSettingsCollection SummerMonthList = new GenericSettingsCollection();
        private GenericSettingsCollection FallMonthList = new GenericSettingsCollection();
        private GenericSettingsCollection WinterMonthList = new GenericSettingsCollection();
        private GenericSettingsCollection SpringMonthList = new GenericSettingsCollection();
        
        public Options()
        {
            InitializeComponent();            
        }

        public Options(UserSettings Usersettings)
        {
            InitializeComponent();
            CollegeOnly = true;
            this.UserSettings = Usersettings;
            this.Colleges = new Colleges();
            TxtXMLDirectory.Text = UserSettings.MyDocuments;
            ShowCollegeOnly();
        }

        public void ShowCollegeOnly()
        {
            // Hide Tabs            
            TabControlOptions.TabPages.RemoveByKey("WebAssign");
            TabControlOptions.TabPages.RemoveByKey("Email");
            TabControlOptions.TabPages.RemoveByKey("Excel");
            TabControlOptions.TabPages.RemoveByKey("AllOthers");
            // Hide Quarter Group
            GrouperQuarter.Visible = false;
            // Hide Outlook
            GrouperOutlook.Visible = false;
            // Hide canel options
            BtnCancel.Visible = false;
            ControlBox = false;
            // Set up the combo box
            SetUpColleges();
        }

        #region Comboboxes Setup

        private void SetUpQuarterComboBoxes()
        {
            #region Summer
            SummerMonthList.Add(6, "June");
            SummerMonthList.Add(7, "July");

            CmbQuarterSummer.DataSource = SummerMonthList;
            CmbQuarterSummer.DisplayMember = "Description";
            CmbQuarterSummer.ValueMember = "Position";
            #endregion

            #region Fall
            FallMonthList.Add(8, "August");
            FallMonthList.Add(9, "September");

            CmbQuarterFall.DataSource = FallMonthList;
            CmbQuarterFall.DisplayMember = "Description";
            CmbQuarterFall.ValueMember = "Position";
            #endregion

            #region Winter
            WinterMonthList.Add(12, "December");
            WinterMonthList.Add(1, "January");

            CmbQuarterWinter.DataSource = WinterMonthList;
            CmbQuarterWinter.DisplayMember = "Description";
            CmbQuarterWinter.ValueMember = "Position";
            #endregion

            #region Spring
            SpringMonthList.Add(3, "March");
            SpringMonthList.Add(4, "April");

            CmbQuarterSpring.DataSource = SpringMonthList;
            CmbQuarterSpring.DisplayMember = "Description";
            CmbQuarterSpring.ValueMember = "Position";
            #endregion
        }


        private void SetUpClickerFormat()
        {
            CmbClickerFormat.DataSource = Clicker.ClickerFormat();
            CmbClickerFormat.DisplayMember = "Description";
            CmbClickerFormat.ValueMember = "Position";
        }

        private void SetUpColleges()
        {
            CmbColleges.DataSource = Colleges;
            CmbColleges.DisplayMember = "DropDownBoxName";
            CmbColleges.ValueMember = "Abbreviation";
        }

        #endregion

        public void SetupEmailUsage()
        {
            ChkEmailUsage.Checked = Properties.Settings.Default.EmailUsage;
            LblEmailNextUsageDate.Text = "Next email will be sent after: " + Properties.Settings.Default.EmailDateAfter.ToShortDateString();
        }

        public void InitilizeExcel(ExcelRollSettings Roll, ExcelClassSettings Class)
        {

            #region Excel Class
            ChkExcelClassExport.Checked = Class.Export;
            ChkExcelClassWaitListExport.Checked = Class.ExportWaitlist;
            if (!Class.TemplateDirectory.EndsWith("\\")) { Class.TemplateDirectory += "\\"; }
            TxtExcelTemplateNameAndDirectory.Text = Class.TemplateDirectory + Class.TemplateFileName;
            TxtExcelSaveAsDirectory.Text = Class.SaveFileDirectory;
            TxtExcelSaveAsFileName.Text = Class.SaveFileName;
            ChkExcelCommonIncludeQuarter.Checked = Class.SaveFileNameIncludesQuarter;
            TxtExcelClassFirstStudentRow.Text = Class.FirstStudent.ToString();
            TxtExcelClassSID.Text = Class.SIDColumnLetter;
            chkExcelClassExportSID.Checked = Class.ExportSID;
            TxtExcelClassSIDLast4.Text = Class.SIDLast4ColumnLetter;
            chkExcelClassExportSIDLast4.Checked = Class.ExportSIDLast4;
            TxtExcelClassLastName.Text = Class.LastNameColumnLetter;
            TxtExcelClassFirstName.Text = Class.FirstNameColumnLetter;
            TxtExcelClassItemCell.Text = Class.ItemCell;
            ChkExcelClassWaitListExport.Checked = Class.ExportWaitlist;

            #region OptHeader Textbox and Checkboxes
            TxtExcelClassOptHeader1.Text = Class.OptHead1ColumnLetter;
            chkExcelClassExportoptHeader1.Checked = Class.ExportoptHead1;

            TxtExcelClassOptHeader2.Text = Class.OptHead2ColumnLetter;
            chkExcelClassExportoptHeader2.Checked = Class.ExportoptHead2;

            TxtExcelClassOptHeader3.Text = Class.OptHead3ColumnLetter;
            chkExcelClassExportoptHeader3.Checked = Class.ExportoptHead3;
            #endregion

            #region OptHeader Labels
            // Change to XML
            LblExcelClassOptHeader1.Text = Class.HeaderNames.Header1;
            chkExcelClassExportoptHeader1.Text = "Export " + Class.HeaderNames.Header1;

            LblExcelClassOptHeader2.Text = Class.HeaderNames.Header2;
            chkExcelClassExportoptHeader2.Text = "Export " + Class.HeaderNames.Header2;

            LblExcelClassOptHeader3.Text = Class.HeaderNames.Header3;
            chkExcelClassExportoptHeader3.Text = "Export " + Class.HeaderNames.Header3;
            #endregion

            #endregion

            #region Excel Roll
            ChkExcelRollExport.Checked = Roll.Export;            
            ChkExcelRollWaitListExport.Checked = Roll.ExportWaitlist;
            ChkExcelLabExport.Checked = Roll.ExportLab;
            txtRollFirstClass.Text = Roll.FirstClass.ToString();
            txtRollClassIncrement.Text = Roll.ClassIncrement.ToString();
            txtRollHeader.Text = Roll.Header;
            txtRollLastName.Text = Roll.LastNameColumnLetter;
            txtRollFirstName.Text = Roll.FirstNameColumnLetter;
            txtRollMondayDateCell.Text = Roll.MondayDateCell;
            txtRollFirstDayCell.Text = Roll.FirstMondayDayCell;
            #endregion

        }

        public void Initilize()
        {
            SetupEmailUsage();

            TxtXMLDirectory.Text = UserSettings.MyDocuments;

            #region Quarter List items
            SetUpQuarterComboBoxes();
            CmbQuarterSummer.SelectedValue = Quarter2Month.SummerMonth;
            CmbQuarterFall.SelectedValue = Quarter2Month.FallMonth;
            CmbQuarterWinter.SelectedValue = Quarter2Month.WinterMonth;
            CmbQuarterSpring.SelectedValue = Quarter2Month.SpringMonth;
            #endregion

            #region User
            TxtOutlookUserEmail.Text = UserSettings.Email;
            ChkUserOverWrite.Checked = UserSettings.OverWriteAll;
            #endregion

            #region Colleges
            SetUpColleges();
            #endregion

            #region Clicker
            TxtClickerDirectory.Text = ClickerSettings.Directory;
            ChkClickerUnderscore.Checked = ClickerSettings.Underscore;
            ChkClickerExport.Checked= ClickerSettings.Export;
            SetUpClickerFormat();
            CmbClickerFormat.SelectedValue = ClickerSettings.SelectedValue;
            #endregion

            #region Wamap
            TxtWamapDirectory.Text = WamapSettings.Directory;
            ChkWamapUnderscore.Checked = WamapSettings.Underscore;
            ChkWamapExport.Checked = WamapSettings.Export;
            #endregion

            #region MTG
            TxtMTGDirectory.Text = MTGSettings.Directory;
            ChkMTGUnderscore.Checked = MTGSettings.Underscore;
            ChkMTGExport.Checked = MTGSettings.Export;
            #endregion

            #region Outlook
            TxtOutlookUserEmail.Text = UserSettings.Email;
            ChkOutlookUnderscore.Checked = MTGSettings.Underscore;
            ChkOutlookUnderscore.Checked = MTGSettings.Export;
            #endregion

            #region WebAssign
            ChkWebAssignExport.Checked = WebAssignSettings.Export;
            TxtWebAssignDirectory.Text = WebAssignSettings.Directory;
            ChkWebAssignUnderscore.Checked = WebAssignSettings.Underscore;
            if (WebAssignSettings.SEPARATOR == WebAssignSeperator.Tab)
            { CmbWebAssignSeperator.Text = "TAB"; }
            else
            { CmbWebAssignSeperator.Text = "COMMA"; }

            if (WebAssignSettings.UserPassword == WebAssignPassword.SID)
            { CmbWebAssignPassword.Text = "SID"; }
            else
            { CmbWebAssignPassword.Text = "LastName"; }

            switch (WebAssignSettings.Username)
            {
                case WebAssignUserName.First_Name_Initial_plus_Last_Name:
                    CmbWebAssignUserName.Text = "First Name Initial + Last Name";
                    break;
                case WebAssignUserName.First_Name_plus_Last_Name:
                    CmbWebAssignUserName.Text = "First Name + Last Name";
                    break;
                case WebAssignUserName.First_Name_Initial_plus_Middle_Initial_plus_Last_Name:
                    CmbWebAssignUserName.Text = "First Name Initial + Middle Initial + Last Name";
                    break;
                default:
                    break;
            }
            #endregion

            InitilizeExcel(ExcelRollSettings, ExcelClassSettings);
        }

        public void SaveValues()
        {
            Properties.Settings.Default.EmailUsage = ChkEmailUsage.Checked;
            UserSettings.CollegeAbbreviation = (string)CmbColleges.SelectedValue;
            UserSettings.MyDocuments = TxtXMLDirectory.Text;

            if (!CollegeOnly)
            {
                UserSettings.OverWriteAll = ChkUserOverWrite.Checked;

                #region Quarter List items
                SetUpQuarterComboBoxes();
                Quarter2Month.SummerMonth = Convert.ToInt32(CmbQuarterSummer.SelectedValue);
                Quarter2Month.FallMonth = Convert.ToInt32(CmbQuarterFall.SelectedValue);
                Quarter2Month.WinterMonth = Convert.ToInt32(CmbQuarterWinter.SelectedValue);
                Quarter2Month.SpringMonth = Convert.ToInt32(CmbQuarterSpring.SelectedValue);
                #endregion

                #region Clicker
                ClickerSettings.Directory = TxtClickerDirectory.Text;
                ClickerSettings.Underscore = ChkClickerUnderscore.Checked;
                ClickerSettings.Export = ChkClickerExport.Checked;
                ClickerSettings.SelectedValue = Convert.ToInt32(CmbClickerFormat.SelectedValue);
                #endregion

                #region Wamap
                WamapSettings.Directory = TxtWamapDirectory.Text;
                WamapSettings.Underscore = ChkWamapUnderscore.Checked;
                WamapSettings.Export = ChkWamapExport.Checked;
                #endregion

                #region MTG
                MTGSettings.Directory = TxtMTGDirectory.Text;
                MTGSettings.Underscore = ChkMTGUnderscore.Checked;
                MTGSettings.Export = ChkMTGExport.Checked;
                #endregion

                #region Outlook
                UserSettings.Email = TxtOutlookUserEmail.Text;
                OutlookSettings.Underscore = ChkOutlookUnderscore.Checked;
                OutlookSettings.Export = ChkOutlookUnderscore.Checked;
                #endregion

                #region WebAssign
                WebAssignSettings.Export = ChkWebAssignExport.Checked;
                WebAssignSettings.Directory = TxtWebAssignDirectory.Text;
                WebAssignSettings.Underscore = ChkWebAssignUnderscore.Checked;

                if (CmbWebAssignSeperator.Text == "TAB")
                { WebAssignSettings.SEPARATOR = WebAssignSeperator.Tab; }
                else
                { WebAssignSettings.SEPARATOR = WebAssignSeperator.Comma; }

                if (CmbWebAssignPassword.Text == "SID")
                { WebAssignSettings.UserPassword = WebAssignPassword.SID; }
                else
                { WebAssignSettings.UserPassword = WebAssignPassword.LastName; }

                switch (CmbWebAssignUserName.Text)
                {
                    case "First Name Initial + Last Name":
                        WebAssignSettings.Username = WebAssignUserName.First_Name_Initial_plus_Last_Name;
                        break;
                    case "First Name + Last Name":
                        WebAssignSettings.Username = WebAssignUserName.First_Name_plus_Last_Name;
                        break;
                    case "First Name Initial + Middle Initial + Last Name":
                        WebAssignSettings.Username = WebAssignUserName.First_Name_Initial_plus_Middle_Initial_plus_Last_Name;
                        break;
                    default:
                        break;
                }

                #endregion

                #region Excel Class

                ExcelClassSettings.TemplateDirectory = Path.GetDirectoryName(TxtExcelTemplateNameAndDirectory.Text);
                ExcelClassSettings.TemplateFileName = Path.GetFileName(TxtExcelTemplateNameAndDirectory.Text);
                ExcelClassSettings.SaveFileDirectory = TxtExcelSaveAsDirectory.Text;
                ExcelClassSettings.SaveFileName = TxtExcelSaveAsFileName.Text;
                ExcelClassSettings.SaveFileNameIncludesQuarter = ChkExcelCommonIncludeQuarter.Checked;
                try
                {
                    ExcelClassSettings.FirstStudent = Convert.ToInt32(TxtExcelClassFirstStudentRow.Text);
                }
                catch
                {
                    ExcelClassSettings.FirstStudent = ExcelClassConfiguration.FirstStudentDefault;
                }
                ExcelClassSettings.SIDColumnLetter = TxtExcelClassSID.Text;
                ExcelClassSettings.Export = ChkExcelClassExport.Checked;
                ExcelClassSettings.ExportWaitlist = ChkExcelClassWaitListExport.Checked;
                ExcelClassSettings.ExportSID = chkExcelClassExportSID.Checked;
                ExcelClassSettings.SIDLast4ColumnLetter = TxtExcelClassSIDLast4.Text;
                ExcelClassSettings.ExportSIDLast4 = chkExcelClassExportSIDLast4.Checked;
                ExcelClassSettings.LastNameColumnLetter = TxtExcelClassLastName.Text;
                ExcelClassSettings.FirstNameColumnLetter = TxtExcelClassFirstName.Text;
                ExcelClassSettings.ItemCell = TxtExcelClassItemCell.Text;

                #region OptHeader Textbox and Checkboxes
                ExcelClassSettings.OptHead1ColumnLetter = TxtExcelClassOptHeader1.Text;
                ExcelClassSettings.ExportoptHead1 = chkExcelClassExportoptHeader1.Checked;

                ExcelClassSettings.OptHead2ColumnLetter = TxtExcelClassOptHeader2.Text;
                ExcelClassSettings.ExportoptHead2 = chkExcelClassExportoptHeader2.Checked;

                ExcelClassSettings.OptHead3ColumnLetter = TxtExcelClassOptHeader3.Text;
                ExcelClassSettings.ExportoptHead3 = chkExcelClassExportoptHeader3.Checked;
                #endregion

                #endregion

                #region Excel Roll
                ExcelRollSettings.Export = ChkExcelRollExport.Checked;
                ExcelRollSettings.ExportWaitlist = ChkExcelRollWaitListExport.Checked;
                try
                {
                    ExcelRollSettings.FirstClass = Convert.ToInt32(txtRollFirstClass.Text);
                }
                catch
                {
                    ExcelRollSettings.FirstClass = ExcelRollConfiguration.FirstClassDefault;
                }

                try
                {
                    ExcelRollSettings.ClassIncrement = Convert.ToInt32(txtRollClassIncrement.Text);
                }
                catch
                {
                    ExcelRollSettings.ClassIncrement = ExcelRollConfiguration.ClassIncrementDefault;
                }
                ExcelRollSettings.Header = txtRollHeader.Text;                
                ExcelRollSettings.LastNameColumnLetter = txtRollLastName.Text;
                ExcelRollSettings.FirstNameColumnLetter = txtRollFirstName.Text;
                ExcelRollSettings.MondayDateCell = txtRollMondayDateCell.Text;
                ExcelRollSettings.FirstMondayDayCell = txtRollFirstDayCell.Text;
                #endregion
            }
        }

        #region Save Cancel Buttons

        private void BtnSave_Click(object sender, EventArgs e)
        {
            SaveValues();
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

        #region Directory Buttons

        private void SetDirectoryText(TextBox textBox)
        {
            string StrFolder = textBox.Text;
            if (!Directory.Exists(StrFolder)) { StrFolder = Properties.Settings.Default.MyDocuments; }

            FolderBrowserDialog FolderBrowserDialog = new FolderBrowserDialog
            {
                SelectedPath = StrFolder
            };

            if (FolderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = FolderBrowserDialog.SelectedPath;
            }
        }

        private void BtnClickerDirectory_Click(object sender, EventArgs e)
        {
            SetDirectoryText(TxtClickerDirectory);
        }

        private void BtnWebAssignDirectory_Click(object sender, EventArgs e)
        {
            SetDirectoryText(TxtWebAssignDirectory);
        }

        private void BtnMTGDirectory_Click(object sender, EventArgs e)
        {
            SetDirectoryText(TxtMTGDirectory);
        }

        private void BtnWamapDirectory_Click(object sender, EventArgs e)
        {
            SetDirectoryText(TxtWamapDirectory);
        }

        private void BtnXMLDirectory_Click(object sender, EventArgs e)
        {
            SetDirectoryText(TxtXMLDirectory);
        }

        private void BtnExcelDirectory_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Workbook (*.xls)|*.xls",
                InitialDirectory = Path.GetDirectoryName(TxtExcelTemplateNameAndDirectory.Text)
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                TxtExcelTemplateNameAndDirectory.Text = openFileDialog.FileName;
            }
        }

        private void BtnFindFile_Click(object sender, EventArgs e)
        {
            SetDirectoryText(TxtExcelSaveAsDirectory);
        }

        private void BtnExcelLoadConfiguration_Click(object sender, EventArgs e)
        {
            Excel Excel = new Excel(UserSettings);

            // make copies
            ExcelRollSettings Roll = ExcelRollSettings.Clone();
            ExcelClassSettings Class =  ExcelClassSettings.Clone();

            Class.TemplateDirectory = Path.GetDirectoryName( TxtExcelTemplateNameAndDirectory.Text);
            Class.TemplateFileName =  Path.GetFileName(TxtExcelTemplateNameAndDirectory.Text);

            if (Excel.LoadConfigurationFromSpreadSheet(Roll,Class))
            {
                // - Found information - set form values
                InitilizeExcel(Roll, Class);
                MessageBox.Show("Retrieved values and updated the form", "Success!");
            }           
        }

        #endregion

        private void CmbColleges_SelectedIndexChanged(object sender, EventArgs e)
        {
            College C = (College)CmbColleges.SelectedItem;
            LbLEmployeeIDpromptvalue.Text = C.EmployeeIDprompt;
            LblWebAssignInstitutionvalue.Text = C.WebAssignInstitution;
            LblRequestURLvalue.Text = C.RequestURL;
        }

        private void LLEmail_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string body = "Hi%20Mike,%0D%0A%0D%0AMy%20college%20information%20is%20as%20follows:%0D%0A%0D%0ACollege%20Full%20Name:%0D%0A%0D%0ACollege%20Abbreviation%0D%0A%0D%0AEmployeeID%20prompt%0D%0A%0D%0AWebAssign%20Institution%20code%0D%0A%0D%0AInstructor%20Briefcase%20RequestURL:%0D%0A%0D%0AThanks,%0D%0A%0D%0A(Your%20Name%20Here)";

            System.Diagnostics.Process.Start("mailto:mpjenck@gmail.com?subject=Instructor%20Briefcase%20Extractor%20-%20My%20College%20is%20missing&body="+body);
        }

    }
}
