
namespace InstructorBriefcaseExtractor.Model
{
    public enum WebAssignPassword { LastName, SID };
    public enum WebAssignSeperator { Tab, Comma };
    public enum WebAssignUserName { First_Name_Initial_plus_Last_Name, First_Name_plus_Last_Name, First_Name_Initial_plus_Middle_Initial_plus_Last_Name };

    public class WebAssignSettings
    {
        private readonly MyDirectory MyDirectory;
        public WebAssignSettings()
        {
            MyDirectory = new MyDirectory();
            Institution = "";
            Username = WebAssignUserName.First_Name_Initial_plus_Last_Name;
            SEPARATOR = WebAssignSeperator.Tab;
            UserPassword = WebAssignPassword.LastName;
        }

        public string Institution { get; set; }
        public bool Underscore { get; set; }
        public bool Export { get; set; }
        public bool ExportWaitlist { get; set; }
        public string Directory
        {
            get
            { return MyDirectory.Directory; }
            set
            { MyDirectory.Directory = value; }
        }

        private WebAssignUserName WebAssignUserName;
        public WebAssignUserName Username
        {
            get { return WebAssignUserName; }
            set { WebAssignUserName = value; }
        }

        private WebAssignSeperator WebAssignSeperator;
        public WebAssignSeperator SEPARATOR
        {
            get { return WebAssignSeperator; }
            set { WebAssignSeperator = value; }
        }

        private WebAssignPassword WebAssignPassword;
        public WebAssignPassword UserPassword
        {
            get { return WebAssignPassword; }
            set { WebAssignPassword = value; }
        }
    }
}
