
namespace InstructorBriefcaseExtractor.Model
{
    public class MTGSettings
    {
        private readonly MyDirectory MyDirectory;
        public MTGSettings()
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
    }
}
