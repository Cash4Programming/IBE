using System;
using System.Globalization;
using System.IO;

using InstructorBriefcaseExtractor.Model; 
using InstructorBriefcaseExtractor.DAL;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    public class Wizard
    {
        #region Events

        public event InstructorBriefcaseExtractor.Model.InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, EventArgs e) => MessageResults?.Invoke(sender, e);
        //protected virtual void SendMessage(object sender, EventArgs e)
        //{
        //    InstructorBriefcaseExtractor.Model.InformationalMessage handler = MessageResults;
        //    if (handler != null)
        //    {
        //        handler(sender, e);
        //    }
        //}

        public event FileExist FileExiStrequest;
        protected virtual void FileExistsMessage(object sender, EventArgs e)
        {
            FileExiStrequest?.Invoke(sender, e);
        }

        //protected virtual void FileExistsMessage(object sender, EventArgs e)
        //{
        //    FileExist handler = FileExiStrequest;
        //    if (handler != null)
        //    {
        //        handler(sender, e);
        //    }
        //}

        #endregion

        #region Properties and Variables
        private bool ExportSuccessful = false;
        public College SelectedCollege { get; set; }
        public UserSettings UserSettings { get; set; }
        public ClickerSettings ClickerSettings { get; set; }
        public MTGSettings MTGSettings { get; set; }
        public OutlookSettings OutlookSettings { get; set; }
        public WamapSettings WamapSettings { get; set; }
        public WebAssignSettings WebAssignSettings { get; set; }
        public ExcelRollSettings ExcelRollSettings { get; set; }
        public ExcelClassSettings ExcelClassSettings { get; set; }
        public Quarter SelectedQuarter { get; set; }
        public Colleges Colleges { get; set; }
        public Quarter2MonthSettings Quarter2Month { get; set; }

        private readonly Quarters myQuarters;               // Available Quarter
        public Quarters Quarters
        {
            get { return myQuarters; }
        }

        private readonly Courses myCourses;                 // Available Courses for selected Quarter
        public Courses Courses
        {
            get { return myCourses; }
        }

        public enum WizardStep { Creation, VerifyTemplates, QuartersFound, ClassesRetrieved, Output, CloseProgram };
        public  WizardStep CurrentCompletedStep;          // Current Completed step of the wizard

        public WizardStep CompletedStep
        {
            get { return CurrentCompletedStep; }
        }

        #endregion

        private void ReadFromXML()
        {
            UserConfiguration UserConfiguration = new UserConfiguration(UserSettings);
            this.UserSettings = UserConfiguration.Load();

            ClickerConfiguration ClickerConfiguration = new ClickerConfiguration(UserSettings);
            this.ClickerSettings = ClickerConfiguration.Load();

            ExcelClassConfiguration ExcelClassConfiguration = new ExcelClassConfiguration(UserSettings);
            this.ExcelClassSettings = ExcelClassConfiguration.Load();

            ExcelRollConfiguration ExcelRollConfiguration = new ExcelRollConfiguration(UserSettings);
            this.ExcelRollSettings = ExcelRollConfiguration.Load();

            MTGConfiguration MTGConfiguration = new MTGConfiguration(UserSettings);
            this.MTGSettings = MTGConfiguration.Load();

            OutlookConfiguration OutlookConfiguration = new OutlookConfiguration(UserSettings);
            this.OutlookSettings = OutlookConfiguration.Load();

            WebAssignConfiguration WebAssignConfiguration = new WebAssignConfiguration(UserSettings);
            this.WebAssignSettings = WebAssignConfiguration.Load();

            WamapConfiguration WamapConfiguration = new WamapConfiguration(UserSettings);
            this.WamapSettings = WamapConfiguration.Load();

            Quarter2MonthConfiguration MyQuarters = new Quarter2MonthConfiguration(UserSettings);
            this.Quarter2Month = MyQuarters.Load();
        }

        public void WriteToXML()
        {
            UserConfiguration userConfiguration= new UserConfiguration(UserSettings);
            userConfiguration.Save();

            ClickerConfiguration ClickerConfiguration = new ClickerConfiguration(UserSettings);
            ClickerConfiguration.Save(this.ClickerSettings);

            ExcelClassConfiguration ExcelClassConfiguration = new ExcelClassConfiguration(UserSettings);
            ExcelClassConfiguration.Save(this.ExcelClassSettings);

            ExcelRollConfiguration ExcelRollConfiguration = new ExcelRollConfiguration(UserSettings);
            ExcelRollConfiguration.Save(this.ExcelRollSettings);

            MTGConfiguration MTGConfiguration = new MTGConfiguration(UserSettings);
            MTGConfiguration.Save(this.MTGSettings);

            OutlookConfiguration OutlookConfiguration = new OutlookConfiguration(UserSettings);
            OutlookConfiguration.Save(this.OutlookSettings);

            WebAssignConfiguration WebAssignConfiguration = new WebAssignConfiguration(UserSettings);
            WebAssignConfiguration.Save(this.WebAssignSettings);

            WamapConfiguration WamapConfiguration = new WamapConfiguration(UserSettings);
            WamapConfiguration.Save(this.WamapSettings);

            Quarter2MonthConfiguration MyQuarters = new Quarter2MonthConfiguration(UserSettings);
            MyQuarters.Save(this.Quarter2Month);
        }

        public string CurrentQuarter()
        {
            Quarter2MonthConfiguration MyQuarters = new Quarter2MonthConfiguration(UserSettings);
            return MyQuarters.CurrentQuarter(Quarter2Month);
        }

        public void GetCollege()
        {
            Options options = new Options(UserSettings);

            options.ShowDialog();
            options.Dispose();
        }

        public void WriteUserSettingsToXML()
        {
            UserConfiguration userConfiguration = new UserConfiguration(UserSettings);
            userConfiguration.Save();
        }

        // Step 1 = get saved User information
        public Wizard()
        {
            this.SelectedQuarter = new Quarter();
            this.myQuarters = new Quarters();
            this.myCourses = new Courses();
            this.UserSettings = new UserSettings();

            if (!File.Exists(UserSettings.PathandFileName))
            {
                if (UserSettings.Domain.ToLower(CultureInfo.CurrentCulture) == "yvcc")
                {
                    UserSettings.CollegeAbbreviation = YVC.Key;
                }
                else
                {
                    GetCollege();
                }
                WriteUserSettingsToXML(); // Write us Version does not exist                
            }

            // get the XML file version
            XMLFileVersionConfiguration XMLFileVersionConfiguration = new XMLFileVersionConfiguration(UserSettings);
            UserSettings.Version = XMLFileVersionConfiguration.CheckVersion();
            if (UserSettings.Version == "") { UserSettings.Version = "1"; }

            // Load All configuration data
            //UserConfiguration UserConfiguration = new UserConfiguration(UserSettings);
            ReadFromXML();

            // Select College
            Colleges = new Colleges();
            this.SelectedCollege = Colleges.FindCollege(UserSettings.CollegeAbbreviation);

            if (SelectedCollege == null)
            {
                if (UserSettings.Domain.ToLower() == "yvcc")
                {
                    UserSettings.CollegeAbbreviation = YVC.Key;
                }
                else
                {
                    GetCollege();
                }

                this.SelectedCollege = Colleges.FindCollege(UserSettings.CollegeAbbreviation);
                this.WebAssignSettings.Institution = SelectedCollege.WebAssignInstitution;
            }

            // do I need to upgrade the xml file?
            if (UserSettings.Version == "1")
            {
                UserSettings.Version = "2";
                // upgrade it
                File.Delete(UserSettings.PathandFileName);

                // Adjust the Filename as they currently contain the Path and Filename
                ExcelClassSettings.SaveFileName = Path.GetFileName(ExcelClassSettings.SaveFileName);
                ExcelClassSettings.TemplateFileName = Path.GetFileName(ExcelClassSettings.TemplateFileName);
                
                WriteToXML();
            }


            //debugging code
            //Courses.AddTestData();
            //ExcelClassSettings.Export = true;
            //ExcelRollSettings.Export = false;
            //ExcelRollSettings.ExportLab = false;
            //ExportExcel();

            CurrentCompletedStep = WizardStep.Creation;
        }

        // Step 2 = Verify Templates
        public void VerifyTemplates()
        {
            string TemplatePath = System.Windows.Forms.Application.ExecutablePath;
            int LastSlash = TemplatePath.LastIndexOf('\\');
            if (LastSlash != -1)
            {
                TemplatePath = TemplatePath.Substring(0, LastSlash + 1) + "Templates\\";
            }
            else
            {
                SendMessage(this, new Information("TemplatePath = " + TemplatePath));
                TemplatePath = "Error"; // Error
            }
            SendMessage(this, new Information("TemplatePath = " + TemplatePath));

            try
            {
                if (!File.Exists(UserSettings.MyDocuments + "Gradebook_Group_33_Template_Percent.xls"))
                {
                    File.Copy(TemplatePath + "Gradebook_Group_33_Template_Percent.xls", UserSettings.MyDocuments + "Gradebook_Group_33_Template_Percent.xls");
                }
                else
                {
                    SendMessage(this, new Information("Gradebook_Group_33_Template_Percent.xls - has been found"));
                }
            }
            catch
            {
                SendMessage(this, new Information("Could not find " + TemplatePath + "Gradebook_Group_33_Template_Percent.xls"));
            }

            try
            {
                if (!File.Exists(UserSettings.MyDocuments + "Gradebook_Group_36_Template_Percent.xls"))
                {
                    File.Copy(TemplatePath + "Gradebook_Group_36_Template_Percent.xls", UserSettings.MyDocuments + "Gradebook_Group_36_Template_Percent.xls");
                }
                else
                {
                    SendMessage(this, new Information("Gradebook_Group_36_Template_Percent.xls - has been found"));
                }
            }
            catch
            {
                SendMessage(this, new Information("Could not find " + TemplatePath + "Gradebook_Group_36_Template_Percent.xls"));
            }
            CurrentCompletedStep = WizardStep.VerifyTemplates;
        }

        // Step 3 = Retrieve All Quarters
        private void GetQuarters()
        {
            try
            {
                // Listen for events
                myQuarters.MessageResults += new InformationalMessage(SendMessage);
                myQuarters.AddQuartersByDate();
                CurrentCompletedStep = WizardStep.QuartersFound;
            }
            catch 
            {
                throw;
            }
            finally
            {
                myQuarters.MessageResults -= new InformationalMessage(SendMessage);
            }
        }

        private void MyQuarters_MessageResults(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        // Step 4 = Retrieve All Courses
        private int GetCourses(CISLogin MyLogin)
        {
            // get the Action URL by retreiveing the login webpage and parsing
            // the page to find the POST Action= item.
            CISAction CISAction = new CISAction();
            string ActionUrl = CISAction.PostURL(SelectedCollege.RequestURL);

            try
            {
                // Listen for events
                myCourses.MessageResults += new InformationalMessage(SendMessage);
                SendMessage(this, new Information("Requesting course(s) from Website."));
                try
                {
                    myCourses.AddFromCIS(MyLogin, ExcelClassSettings, ActionUrl);
                }
                catch (Exception ex)
                {
                    if (myCourses.Count() == 0)
                    {
                        string temp = ex.Message;
                        throw;
                    }
                        // otherwise continue to process what was found
                }
                
                
                // Are there any courses?
                if (myCourses.Count() > 0)
                {
                    CurrentCompletedStep = WizardStep.ClassesRetrieved;
                    SendMessage(this, new Information(myCourses.Count()+ " course(s) retreived."));
                    try
                    {
                        // set Monday Date
                        ExcelRollSettings.FirstDay= Convert.ToDateTime(myCourses[0].StartDate);
                    }
                    catch
                    { }
                    
                }                              
            }
            catch
            {
                throw;
            }
            finally
            {
                // release events
                myCourses.MessageResults -= new InformationalMessage(SendMessage);
            }
            return myCourses.Count();
        }
    
        // Step 4 = Output select information
        private void ExportClicker()
        {
            if (ClickerSettings.Export)
            {
                Clicker Clicker = new Clicker();
                try
                {
                    Clicker.MessageResults += new InformationalMessage(MessageResults);
                    Clicker.Export(ClickerSettings, UserSettings, myCourses);
                }
                catch //(Exception Ex)
                {
                    throw;
                }
                finally
                {
                    Clicker.MessageResults -= new InformationalMessage(MessageResults);
                }
            }
        }
        private void ExportExcel()
        {
            if ((ExcelClassSettings.Export)||(ExcelRollSettings.Export) || (ExcelRollSettings.ExportLab))
            {
                Excel Excel = new Excel(UserSettings);
                
                try
                {
                    Excel.MessageResults += new InformationalMessage(MessageResults);
                    Excel.FileExiStrequest += new FileExist(FileExistsMessage);
                    Excel.Export(ExcelClassSettings, ExcelRollSettings, myCourses, SelectedQuarter);
                }
                catch
                {
                    throw;
                }
                finally
                {
                    Excel.MessageResults -= new InformationalMessage(MessageResults);
                    Excel.FileExiStrequest -= new FileExist(FileExistsMessage);
                }
            }
        }
        private void ExportMTG()
        {
            if (MTGSettings.Export)
            {
                MTG MTG = new MTG();
                try
                {
                    MTG.MessageResults += new InformationalMessage(MessageResults);
                    MTG.Export(MTGSettings, UserSettings, myCourses);
                }
                catch //(Exception Ex)
                {
                    throw;
                }
                finally
                {
                    MTG.MessageResults -= new InformationalMessage(MessageResults);
                }                
            }
        }
        private void ExportOutlook()
        {
            if (OutlookSettings.Export)
            {
                Outlook Out = new Outlook();
                try
                {
                    Out.MessageResults += new InformationalMessage(MessageResults);
                    Out.Export(OutlookSettings, myCourses);
                }
                catch //(Exception Ex)
                {
                    throw;
                }
                finally
                {
                    Out.MessageResults -= new InformationalMessage(MessageResults);
                }
                
            }
        }
        private void ExportWamap()
        {
            if (WamapSettings.Export)
            {
                Wamap Wamap = new Wamap();                
                try
                {
                    Wamap.MessageResults += new InformationalMessage(MessageResults);
                    Wamap.FileExiStrequest += new FileExist(FileExistsMessage);
                    Wamap.Export(WamapSettings, UserSettings, myCourses);
                }
                catch //(Exception Ex)
                {
                    throw;
                }
                finally
                {
                    Wamap.MessageResults -= new InformationalMessage(MessageResults);
                    Wamap.FileExiStrequest -= new FileExist(FileExistsMessage);
                }
            }
        }
        private void ExportWebAssign()
        {
            if (WebAssignSettings.Export)
            {
                WebAssign WebAssign = new WebAssign();
                try
                {
                    WebAssign.MessageResults += new InformationalMessage(MessageResults);
                    WebAssign.FileExiStrequest += new FileExist(FileExistsMessage);
                    WebAssign.Export(WebAssignSettings, UserSettings, myCourses);
                }
                catch //(Exception Ex)
                {
                    throw;
                }
                finally
                {
                    WebAssign.MessageResults -= new InformationalMessage(MessageResults);
                    WebAssign.FileExiStrequest -= new FileExist(FileExistsMessage);
                }
            }
        }

        public void Email_Usage()
        {
            if (ExportSuccessful)
            {
                Outlook MyOutlook = new Outlook();

                string Body = "Dear User,\r\n\r\nPlease send this email to me so I can keep track of usage.\r\n\r\nThank you very much.\r\n\r\nMike Jenck\r\n\r\n";

                Body += "From Username: " + UserSettings.UserName + "\r\nEmail: " + UserSettings.Email + "\r\n";
                if (ClickerSettings.Export) { Body += "Clicker is selected for export.\r\n"; }
                if (OutlookSettings.Export) { Body += "Outlook is selected for export.\r\n"; }
                if (WamapSettings.Export) { Body += "Wamap is selected for export.\r\n"; }
                if (WebAssignSettings.Export) { Body += "WebAssign is selected for export.\r\n"; }
                if (ExcelClassSettings.Export) { Body += "ExcelClass is selected for export.\r\n"; }
                if (ExcelRollSettings.Export) { Body += "ExcelRoll is selected for export.\r\n"; }
                if (MTGSettings.Export) { Body += "MTG is selected for export.\r\n"; }
                string Subject = "IBE Usage Report";

                if (Properties.Settings.Default.EmailUsage)
                {
                    DateTime NextEmailDate = Properties.Settings.Default.EmailDateAfter;
                    if (DateTime.Now >= NextEmailDate)
                    {
                        Properties.Settings.Default.EmailDateAfter = NextEmailDate.AddDays(60);
                        Properties.Settings.Default.Save();
                        Body += "\r\n\r\nNext email date will be on ar after: " + Properties.Settings.Default.EmailDateAfter.ToShortDateString() + ".";
                        if (MyOutlook != null)
                        {
                            bool Successful = MyOutlook.CreateEmail(Properties.Settings.Default.EmailTo, Subject, Body);
                            if (!Successful)
                            {
                                // Fall back if outlook does not work
                                Mail.EmailDeveloper(Properties.Settings.Default.EmailTo, Subject, Body);
                            }
                        }
                        else
                        {
                            Mail.EmailDeveloper(Properties.Settings.Default.EmailTo, Subject, Body);
                        }

                        Properties.Settings.Default.Save();
                    }
                }
            }
        }


        // Notes: Completion marks
        // Step 1 - upon successful creation of this class
        // Step 2 - upon successful Verifaction of excel templates
        // Step 3 - upon successful retreval of YRQ (quarter) information
        // Step 4 - upon successful retreval Courses for the selected YRQ
        // Step 5 - upon successful outputting of desired formats
        public bool StepNext()
        {
            bool retval = true;
            switch (CurrentCompletedStep)
            {
                case WizardStep.Creation:
                    VerifyTemplates();
                    break;

                case WizardStep.VerifyTemplates:
                    // Get all available quarters
                    GetQuarters();
                    CurrentCompletedStep = WizardStep.QuartersFound;                    
                    break;
                    
                case WizardStep.QuartersFound:
                    // Get courses for selected quarters

                    CISLogin myLogin = new CISLogin(UserSettings, SelectedQuarter);
                    
                    // retreive the courses
                    int numberofcourses = GetCourses(myLogin);
                    if (numberofcourses > 0)
                    {
                        SendMessage(this, new Information(numberofcourses.ToString(CultureInfo.CurrentCulture)+" classes retrieved!"));
                        CurrentCompletedStep = WizardStep.ClassesRetrieved;
                    }
                    else
                    {
                        retval = false;
                        SendMessage(this, new Information("You are not teaching any courses!"));
                    }
                    break;
                case WizardStep.ClassesRetrieved:
                    try
                    {
                        ExportClicker();
                        ExportMTG();
                        ExportWamap();
                        ExportWebAssign();                        
                        ExportOutlook();
                        ExportExcel();
                        ExportSuccessful = true;
    }
                    catch
                    {
                        throw;
                    }
                    finally
                    {
                        CurrentCompletedStep = WizardStep.Output; 
                    }                    
                    break;
                case WizardStep.Output:
                    Email_Usage();

                    CurrentCompletedStep = WizardStep.CloseProgram;
                    break;
                default:                    
                    break;
            }

            return retval;
        }

        public void StepBack()
        {
            switch (CurrentCompletedStep)
            {
                case WizardStep.Creation:
                    // Can't step back
                    break;
                case WizardStep.QuartersFound:
                    // Can't step back
                    break;
                case WizardStep.ClassesRetrieved:
                    // Go back to Quarters fond and have user select the desired quarter
                    // to Retrieve
                    CurrentCompletedStep = WizardStep.QuartersFound;
                    break;
                case WizardStep.Output:
                    // Go back to the Classes Retrieved and have user select classes for output
                    CurrentCompletedStep = WizardStep.ClassesRetrieved;
                    break;
                default:
                    break;
            }
        }
    }
}
