
namespace InstructorBriefcaseExtractor.Model
{
    public class Pointer
    {
        public string ItemNumber { get; set; }
        public string NewName { get; set; }
        public int WorksheetIndex { get; set; }
        public int CourseIndex { get; set; }
    }

    public class ExcelLink
    {
        private readonly int MaxNumberOfCourses;
        public ExcelLink(int CourseCount)
        {
            MaxNumberOfCourses = CourseCount;
            GetPointers = new Pointer[MaxNumberOfCourses];
        }
        public string ClassName { get; set; }
        public Pointer[] GetPointers;

        public void AddToPointer(string ItemNumber, int CourseIndex, string Newname)
        {
            for (int index = 0; index < MaxNumberOfCourses; index++)
            {
                if (GetPointers[index] == null)
                {
                    GetPointers[index] = new Pointer
                    {
                        ItemNumber = ItemNumber,
                        CourseIndex = CourseIndex,
                        NewName = Newname
                    };
                    break;
                }
            }
        }

        public void Set(string Itemnumber, int WorkSheetIndex)
        {
            for (int i = 0; i < MaxNumberOfCourses; i++)
            {
                if (GetPointers[i] != null)
                {
                    if (GetPointers[i].ItemNumber == Itemnumber)
                    {
                        GetPointers[i].WorksheetIndex = WorkSheetIndex;
                        break;
                    }
                }
            }
        }

        public int Count()
        {
            for (int index = 0; index < MaxNumberOfCourses; index++)
            {
                if (GetPointers[index] == null)
                {
                    return index;
                }
            }
            return MaxNumberOfCourses;
        }

        public string GetItemNumber(int WorkSheetIndex)
        {
            for (int i = 0; i < MaxNumberOfCourses; i++)
            {
                if (GetPointers[i] != null)
                {
                    if (GetPointers[i].WorksheetIndex == WorkSheetIndex)
                    {
                        return GetPointers[i].ItemNumber;
                    }
                }
                else { break; }
            }
            return "";
        }

        public int GetCourseIndex(int WorkSheetIndex)
        {
            for (int i = 0; i < MaxNumberOfCourses; i++)
            {
                if (GetPointers[i] != null)
                {
                    if (GetPointers[i].WorksheetIndex == WorkSheetIndex)
                    {
                        return GetPointers[i].CourseIndex;
                    }
                }
                else { break; }
            }
            return -1;
        }
    }
}
