

namespace InstructorBriefcaseExtractor.Model
{
    public class ClassNameAndItemNumber
    {
        public ClassNameAndItemNumber(string ClassClassNameAndItemNumber)
        {
            _ClassClassNameAndItemNumber = ClassClassNameAndItemNumber;
        }

        private string _ClassClassNameAndItemNumber;
        public string ClassClassNameAndItemNumber
        {
            set { _ClassClassNameAndItemNumber = value; }
            get { return _ClassClassNameAndItemNumber; }
        }
    }
}
