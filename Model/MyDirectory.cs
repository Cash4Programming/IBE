

namespace InstructorBriefcaseExtractor.Model
{
    public class MyDirectory
    {
        string MyDirectoryInternal = "";
        public string Directory
        {
            get
            {
                return MyDirectoryInternal;
            }
            set
            {
                if (value != "")
                {
                    if (value.EndsWith("\\"))
                    {
                        MyDirectoryInternal = value;
                    }
                    else
                    {
                        MyDirectoryInternal = value + "\\";
                    }
                }
                else
                {
                    MyDirectoryInternal = "";
                }

            }
        }
    }
}
