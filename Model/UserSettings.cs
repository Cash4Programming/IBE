using System;
using System.Globalization;
using System.IO;

namespace InstructorBriefcaseExtractor.Model
{
    public class UserSettings
    {
        private readonly MyDirectory MyDocumentDirectory;
        public UserSettings()
        {
            MyDocumentDirectory = new MyDirectory();
            CollegeAbbreviation = "";
            EmployeeID = "";
            EmployeePIN = "";
            Email = "";
            Version = "";

            string domainUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string[] parts = domainUser.Split('\\');
            Domain = parts[0];
            UserName = parts[1].ToLower(CultureInfo.CurrentCulture);
            FileName = "IBE_" + UserName + ".XML";
            MyDocuments = GetMyDocuments();
        }

        public string GetMyDocuments()
        {
            // get current MyDocuments Folder
            MyDocumentDirectory.Directory = Properties.Settings.Default.MyDocuments;
            if (MyDocumentDirectory.Directory == "")
            {
                MyDocumentDirectory.Directory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Properties.Settings.Default.MyDocuments = MyDocumentDirectory.Directory + "Instructor Briefcase Extractor\\";
                Properties.Settings.Default.Save();
            }

            // Make sure the Directory exists
            if (!Directory.Exists(MyDocumentDirectory.Directory)) { Directory.CreateDirectory(MyDocumentDirectory.Directory); }

            return MyDocumentDirectory.Directory;
        }

        public UserSettings Clone()
        {
            UserSettings U = new UserSettings
            {
                CollegeAbbreviation = this.CollegeAbbreviation,
                Domain = this.Domain,
                Email = this.Email,
                EmployeeID = this.EmployeeID,
                EmployeePIN = this.EmployeePIN,
                FileName = this.FileName,
                MyDocuments = this.MyDocuments,
                OverWriteAll = this.OverWriteAll,
                UserName = this.UserName,
                Version = this.Version
            };

            return U;
        }

        // Saved to XML File
        public string CollegeAbbreviation { get; set; }
        public string EmployeeID { get; set; }
        public string EmployeePIN { get; set; }
        public string Email { get; set; }
        public bool OverWriteAll { get; set; }

        // Not Saved as these are loaded/created at runtime
        public string Domain { get; set; }
        public string Version { get; set; }
        public string FileName { get; set; }
        public string UserName { get; set; }
        public string PathandFileName
        {
            get
            {
                return MyDocuments + FileName;
            }
        }        
        public string MyDocuments
        {
            get
            {
                return MyDocumentDirectory.Directory;
            }
            set
            {
                MyDocumentDirectory.Directory = value;
            }
        }
       
    }
}
