
namespace InstructorBriefcaseExtractor.Model
{
    public class MTGSettings
    {
        private MyDirectory MyDirectory;
        public MTGSettings()
        {
            MyDirectory = new MyDirectory();
        }

        public bool Underscore { get; set; }
        public bool Export { get; set; }
        public string Directory
        {
            get
            { return MyDirectory.Directory; }
            set
            { MyDirectory.Directory = value; }
        }
    }
}
