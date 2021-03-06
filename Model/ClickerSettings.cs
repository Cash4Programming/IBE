﻿
namespace InstructorBriefcaseExtractor.Model
{
    public class ClickerSettings
    {
        private readonly MyDirectory MyDirectory;
        public ClickerSettings()
        {
            MyDirectory = new MyDirectory();
        }
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
        public int SelectedValue { get; set; }
    }
}
