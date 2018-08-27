
namespace InstructorBriefcaseExtractor.Model
{
    public class ClickerSettings
    {
        private MyDirectory MyDirectory;
        public ClickerSettings()
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
        public int SelectedValue { get; set; }
    }
}
