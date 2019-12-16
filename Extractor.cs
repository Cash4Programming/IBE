using System;
using System.Drawing;
using System.Windows.Forms;

using InstructorBriefcaseExtractor.BLL;
using InstructorBriefcaseExtractor.Utility;
using InstructorBriefcaseExtractor.Model;


namespace InstructorBriefcaseExtractor
{
    public partial class Extractor : Form
    {
        private readonly static string CrLf = "\r\n";
        
        // 510 x 410 - default size
        private int smallheight = 500;
        private int bigheight = 695;
        private int defaultwidth = 563;
        // 

        private Wizard myWizard;
        private bool blnAddClassesToCheckList;
        private bool Initilizing = false;

        #region Events

        protected virtual void SendMessage(object sender, EventArgs e)
        {
            InstructorBriefcaseExtractor.Model.Information Information = (InstructorBriefcaseExtractor.Model.Information)e;
            RtbOutput.Text = Information.Message + CrLf + RtbOutput.Text;
            RtbOutput.Refresh();

            RtbHTML.Text += Information.Message + CrLf;
            RtbHTML.Refresh();            
        }

        private void FileExistsMessage(object sender, EventArgs e)
        {
            InstructorBriefcaseExtractor.Model.FileInformation FE = (InstructorBriefcaseExtractor.Model.FileInformation)e;
            FileExistsDialog myDialog = new FileExistsDialog(FE);
            myDialog.ShowDialog(this);
            myDialog.Dispose();
        }

        // Function to start and stop listening for events.
        private void EndMessages()
        {
            try
            {
                myWizard.MessageResults -= new InstructorBriefcaseExtractor.Model.InformationalMessage(SendMessage);
                myWizard.FileExiStrequest -= new InstructorBriefcaseExtractor.Model.FileExist(FileExistsMessage);
            }
            catch 
            { }
        }

        private void StartMessages()
        {
            try
            {
                myWizard.MessageResults += new InformationalMessage(SendMessage);
                myWizard.FileExiStrequest += new FileExist(FileExistsMessage);
            }
            catch 
            {
                MessageBox.Show("Unable to listen to application messages!");
            }            
        }

        #endregion

        private string GetVersion()
        {
            //// Get the entry point. 
            //Assembly entryPoint = Assembly.GetEntryAssembly();

            //// Get the name of the assembly. 
            //AssemblyName entryPointName = entryPoint.GetName();

            //// Get the version. 
            //Version entryPointVersion = entryPointName.Version;

            string ver = Application.ProductVersion.ToString();            
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                System.Deployment.Application.ApplicationDeployment ad =
                System.Deployment.Application.ApplicationDeployment.CurrentDeployment;
                ver = ad.CurrentVersion.ToString();
            }

            return ver;
        }

        public Extractor()
        {
            InitializeComponent();
            
            try
            {
                // Set button text and form size
                BtnMoreLess.Text = "More >>";
                this.Size = new Size(defaultwidth, smallheight);
                blnAddClassesToCheckList = true;
                
                // version
                LblVersion.Text = "Version = " + GetVersion();                
                myWizard = new Wizard(); // Create a Wizard object
                StartMessages();                 // listen to events of myWizard (in Events region)
                if (myWizard.CompletedStep == Wizard.WizardStep.Creation)
                { myWizard.StepNext(); }
                if (myWizard.CompletedStep == Wizard.WizardStep.VerifyTemplates)
                { myWizard.StepNext(); }
                SetInitialValues(myWizard.SelectedCollege.RequestURL);              // update all fields and checkboxs on forms

                PostStepUpdateScreen();
                LblEmployeeID.Text = myWizard.SelectedCollege.EmployeeIDprompt;
            }
            catch (Exception Ex)
            {
                Mail.EmailDeveloper(Properties.Settings.Default.EmailTo, "Error - IBE", Ex.Message);
                MessageBox.Show(Ex.Message);
            }            
        }

        #region Set values

        private void SetQuarterPanalValues(string RequestURL)
        {
            Initilizing = true;  // suppress the message window
            TxtAddress.Text = RequestURL;
            TxtUserName.Text = myWizard.UserSettings.EmployeeID;
            TxtPIN.Text = myWizard.UserSettings.EmployeePIN;

            foreach (Quarter Q in myWizard.Quarters)
            {
                CmbQuarterName.Items.Add(Q.Name);
            }

            //
            // if there are no quarter let an error occur
            CmbQuarterName.Text = myWizard.CurrentQuarter();// CmbQuarterName.Items[0].ToString();
            lblYRQ.Text = myWizard.Quarters[CmbQuarterName.Text].YRQ;
        }

        private void SetCoursePanalValues()
        {
            Initilizing = true;  // suppress the message window
            // Courses Panel
            CbkClickerExport.Checked = myWizard.ClickerSettings.Export;
            CbkClickerExportWaitlist.Checked = myWizard.ClickerSettings.ExportWaitlist;

            CbkMTGExport.Checked = myWizard.MTGSettings.Export;
            CbkMTGExportWaitlist.Checked = myWizard.MTGSettings.ExportWaitlist;

            CbkExcelClassExport.Checked = myWizard.ExcelClassSettings.Export;
            CbkExcelClassExportWaitlist.Checked = myWizard.ExcelClassSettings.ExportWaitlist;

            CbkExcelLabExport.Checked = myWizard.ExcelRollSettings.ExportLab;
            CbkExcelRollExport.Checked = myWizard.ExcelRollSettings.Export;            
            CbkExcelRollExportWaitlist.Checked = myWizard.ExcelRollSettings.ExportWaitlist;

            CbkOutlookExport.Checked = myWizard.OutlookSettings.Export;
            CbkOutlookExportWaitlist.Checked = myWizard.OutlookSettings.ExportWaitlist;

            CbkWamapExport.Checked = myWizard.WamapSettings.Export;
            CbkWamapExportWaitlist.Checked = myWizard.WamapSettings.ExportWaitlist;

            CbkWebAssignExport.Checked = myWizard.WebAssignSettings.Export;
            CbkWebAssignExportWaitlist.Checked = myWizard.WebAssignSettings.ExportWaitlist;

            ChkOverWrite.Checked = myWizard.UserSettings.OverWriteAll;
            Initilizing = false;
        }

        private void SetInitialValues(string RequestURL)
        {
            SetQuarterPanalValues(RequestURL);
            SetCoursePanalValues();
        }

        private void AddClassesToCheckedListBox()
        {
            if (blnAddClassesToCheckList)
            {
                blnAddClassesToCheckList = false;
                ClbCourses.Items.Clear();
                LblCourses.Text = "Select " + myWizard.Courses[0].QuarterName + " Course(s) to Export";
                foreach (Course C in myWizard.Courses)
                {
                    int ItemNum = ClbCourses.Items.Add(C.Display);
                    ClbCourses.SetItemChecked(ItemNum, true);
                }
            }
        }

        #endregion

        #region Wizard Pre/Post processing

        private void SaveExportSettings()
        {
            myWizard.ClickerSettings.Export = CbkClickerExport.Checked;
            myWizard.ClickerSettings.ExportWaitlist = CbkClickerExportWaitlist.Checked;

            myWizard.MTGSettings.Export = CbkMTGExport.Checked;
            myWizard.MTGSettings.ExportWaitlist = CbkMTGExportWaitlist.Checked;

            myWizard.ExcelClassSettings.Export = CbkExcelClassExport.Checked;
            myWizard.ExcelClassSettings.ExportWaitlist = CbkExcelClassExportWaitlist.Checked;

            myWizard.ExcelRollSettings.Export = CbkExcelRollExport.Checked;
            myWizard.ExcelRollSettings.ExportLab = CbkExcelLabExport.Checked;
            myWizard.ExcelRollSettings.ExportWaitlist = CbkExcelRollExportWaitlist.Checked;

            myWizard.OutlookSettings.Export = CbkOutlookExport.Checked;
            myWizard.OutlookSettings.ExportWaitlist = CbkOutlookExportWaitlist.Checked;

            myWizard.WebAssignSettings.Export = CbkWebAssignExport.Checked;
            myWizard.WebAssignSettings.ExportWaitlist = CbkWebAssignExportWaitlist.Checked;

            myWizard.WamapSettings.Export = CbkWamapExport.Checked;
            myWizard.WamapSettings.ExportWaitlist = CbkWamapExportWaitlist.Checked;

            myWizard.UserSettings.OverWriteAll = ChkOverWrite.Checked;
        }
                
        private void BackPreStepCleanUp()
        {
            switch (myWizard.CompletedStep)
            {
                case Wizard.WizardStep.QuartersFound:
                    break;
                case Wizard.WizardStep.ClassesRetrieved:
                    // Save user information to UserInfo
                    // Courses Panel
                    SaveExportSettings();

                    break;
                case Wizard.WizardStep.Output:
                    break;
                case Wizard.WizardStep.CloseProgram:
                    break;
                default:
                    break;
            }
        }

        // save any user selected Options
        private void NextPreStepCleanUp()
        {
            switch (myWizard.CompletedStep)
            {
                case Wizard.WizardStep.QuartersFound:
                    // Save user information to UserInfo
                    myWizard.UserSettings.EmployeeID = TxtUserName.Text;
                    myWizard.UserSettings.EmployeePIN = TxtPIN.Text;
                    myWizard.WriteUserSettingsToXML();

                    myWizard.SelectedQuarter.Name = CmbQuarterName.Text;
                    myWizard.SelectedQuarter.YRQ = myWizard.Quarters[CmbQuarterName.Text].YRQ;

                    // Make sure that there are no classes
                    myWizard.Courses.Clear();

                    // Now go and retrieve classes
                    blnAddClassesToCheckList = true;  // set flag for courses

                    break;
                case Wizard.WizardStep.ClassesRetrieved:
                    // Save user information to UserInfo
                    // Courses Panel
                    SaveExportSettings();
                    myWizard.WriteUserSettingsToXML();

                    // Set all course NOT to export
                    myWizard.Courses.SetAllExport(false);

                    // now set only the one the user selected to export
                    for (int i = 0; i < ClbCourses.CheckedItems.Count; i++)
                    {
                        string CourseIndex = ClbCourses.CheckedItems[i].ToString().Substring(0, 4);
                        myWizard.Courses[CourseIndex].Export = true;
                    }


                    // Show user results of output panel before stepping
                    ShowOutputPanel();

                    // Clear Results box
                    RtbOutput.Text = "";
                    break;
                case Wizard.WizardStep.Output:
                    break;
                case Wizard.WizardStep.CloseProgram:
                    break;
                default:
                    break;
            }
        }

        // update the screen to reflect the changes 
        private void PostStepUpdateScreen()
        {
            switch (myWizard.CompletedStep)
            {
                case Wizard.WizardStep.Creation:
                    // Can't step back
                    HideAllPanels();
                    BtnBack.Enabled = false;
                    BtnNext.Enabled = true;
                    BtnNext.Text = "Next";
                    break;
                case Wizard.WizardStep.QuartersFound:
                    // Have user select the desired quarter
                    BtnBack.Enabled = false;
                    BtnNext.Enabled = true;
                    BtnNext.Text = "Next";
                    ShowQuarterPanel();
                    break;
                case Wizard.WizardStep.ClassesRetrieved:
                    // Have user select the desired courses for output
                    BtnBack.Enabled = true;
                    BtnNext.Enabled = true;
                    BtnNext.Text = "Next";
                    BtnCancel.Enabled = true;
                    AddClassesToCheckedListBox();
                    ShowCoursePanel();
                    break;
                case Wizard.WizardStep.Output:
                    // Show user results of output
                    BtnBack.Enabled = true;
                    BtnNext.Enabled = true;

                    ShowOutputPanel();
                    break;
                case Wizard.WizardStep.CloseProgram:
                    EndMessages();  // unhook events
                    this.Close();   // Close
                    break;
            }
        }

        #endregion

        #region Panels

        private void HideAllPanels()
        {            
            PanelQuarterFound.Visible = false;
            PanelCourseRetrieved.Visible = false;
            PanelOutput.Visible = false;
        }

        private void ShowQuarterPanel()
        {
            PanelQuarterFound.Left = 10;
            PanelQuarterFound.Top = 25;
            PanelQuarterFound.Visible = true;

            PanelCourseRetrieved.Visible = false;
            PanelOutput.Visible = false;
            blnAddClassesToCheckList = true;
        }

        private void ShowCoursePanel()
        {            
            PanelQuarterFound.Visible = false;
            PanelCourseRetrieved.Visible = true;
            PanelOutput.Visible = false;
        }

        private void ShowOutputPanel()
        {            
            PanelQuarterFound.Visible = false;
            PanelCourseRetrieved.Visible = false;
            PanelOutput.Left = 10;
            PanelOutput.Top = 25;
            PanelOutput.Visible = true;
            // Set buttons
            BtnNext.Text = "Finish";
            BtnCancel.Enabled = false;
        }

        #endregion

        #region Checkbox/Combo box change events

        private void CheckForExportOptions()
        {
            bool ExportType = CbkClickerExport.Checked ||
                              CbkOutlookExport.Checked ||
                              CbkMTGExport.Checked ||
                              CbkExcelClassExport.Checked ||
                              CbkExcelRollExport.Checked ||
                              CbkExcelLabExport.Checked ||
                              CbkWamapExport.Checked ||
                              CbkWamapExport.Checked ||
                              CbkWebAssignExport.Checked;
            if (!Initilizing)
            {
                if ((ClbCourses.CheckedItems.Count == 0) || !ExportType)
                {
                    BtnNext.Enabled = false;
                    MessageBox.Show("No course to output or no output Option selected");
                }
                else
                {
                    BtnNext.Enabled = true;
                }
            }
        }

        private void ClbCourses_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckForExportOptions();   
        }

        private void CbkOutlookExport_CheckedChanged(object sender, EventArgs e)
        {
            CheckForExportOptions(); 
        }

        private void CbkMTGExport_CheckedChanged(object sender, EventArgs e)
        {
            CheckForExportOptions(); 
        }

        private void CbkWebAssignExport_CheckedChanged(object sender, EventArgs e)
        {
            CheckForExportOptions(); 
        }

        private void CbkClickerExport_CheckedChanged(object sender, EventArgs e)
        {
            CheckForExportOptions(); 
        }

        private void CbkWamapExport_CheckedChanged(object sender, EventArgs e)
        {
            CheckForExportOptions(); 
        }

        private void CbkExcelClassExport_CheckedChanged(object sender, EventArgs e)
        {
            CheckForExportOptions();
        }

        private void CbkExcelRollExport_CheckedChanged(object sender, EventArgs e)
        {
            CheckForExportOptions();
        }

        #endregion

        #region Buttons/Menu Selection

        private void BtnNext_Click(object sender, EventArgs e)
        {
            try
            {                
                BtnNext.Enabled = false;
                NextPreStepCleanUp();
                if (myWizard.StepNext())
                {
                    PostStepUpdateScreen();
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message,"Next button error");
            }
            finally
            {
                if (myWizard.CompletedStep == Wizard.WizardStep.ClassesRetrieved)
                {
                    CheckForExportOptions();
                }
                else
                {
                    BtnNext.Enabled = true;
                }
            }
        }

        private void BtnBack_Click(object sender, EventArgs e)
        {
            try
            {
                BackPreStepCleanUp();
                myWizard.StepBack();
                PostStepUpdateScreen();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Back button error");
            }
        }

        private void BtnMoreLess_Click(object sender, EventArgs e)
        {
            if (BtnMoreLess.Text == "More >>")
            {
                BtnMoreLess.Text = "<< Less";
                this.Size = new Size(defaultwidth, bigheight);
            }
            else
            {
                BtnMoreLess.Text = "More >>";
                this.Size = new Size(defaultwidth, smallheight);
            }
        }

        private void FormCloseCleanup()
        {
            EndMessages();     // unhook events
            NextPreStepCleanUp();  // Used to save any changes by the user to disk
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            FormCloseCleanup();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCloseCleanup();
        }

        #endregion

        #region Options

        public void ShowOptions()
        {
            SaveExportSettings();
            Options options = new Options
            {
                // set Form properties with a link to the Wizard variablles
                UserSettings = myWizard.UserSettings,
                Colleges = myWizard.Colleges,
                ClickerSettings = myWizard.ClickerSettings,
                MTGSettings = myWizard.MTGSettings,
                OutlookSettings = myWizard.OutlookSettings,
                WamapSettings = myWizard.WamapSettings,
                WebAssignSettings = myWizard.WebAssignSettings,
                ExcelRollSettings = myWizard.ExcelRollSettings,
                ExcelClassSettings = myWizard.ExcelClassSettings,
                Quarter2Month = myWizard.Quarter2Month
            };

            options.Initilize();
            
            // get form values
            if (options.ShowDialog() == DialogResult.OK)
            {
                myWizard.WriteToXML();
                SetCoursePanalValues();
            }
            options.Dispose();
        }

        private void ConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowOptions();
        }

        #endregion

        private void Extractor_FormClosed(object sender, FormClosedEventArgs e)
        {
            myWizard.WriteToXML();                      
        }

        private void CmbQuarterName_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblYRQ.Text = myWizard.Quarters[CmbQuarterName.Text].YRQ;
        }
    }
}