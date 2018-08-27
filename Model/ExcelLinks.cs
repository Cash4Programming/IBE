using System;
using System.Collections;

namespace InstructorBriefcaseExtractor.Model
{
    public class ExcelLinks : CollectionBase
    {
        private readonly int MaxNumberOfCourses;
        public ExcelLinks(int CourseCount)
        {
            MaxNumberOfCourses = CourseCount;
        }

        public ExcelLink Add(string Name, String ClassItemNumber,int CourseIndex)
        {
            ExcelLink value = this.GetExcelLink(Name);
            if (value == null)
            {
                // doesn't exit - create it
                value = new ExcelLink(MaxNumberOfCourses)
                {
                    ClassName = Name
                };
                List.Add(value);                
            }
            string NewName = Name + " (" + ClassItemNumber + ")";
            // now add the information to the pointer
            value.AddToPointer(ClassItemNumber, CourseIndex, NewName);

            return value;
        }

        public new void Clear()
        {
            this.List.Clear();
        }

        public new int Count()
        {
            return this.List.Count;
        }

        public ExcelLink GetExcelNewNameLink(string NewName)
        {
            foreach (ExcelLink myVar in this.List)
            {
                for (int i = 0; i < myVar.GetPointers.Length; i++)
                {
                    if (myVar.GetPointers[i] != null)
                    {
                        if (myVar.GetPointers[i].NewName.ToLower() == NewName.ToLower())
                        {
                            return myVar;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }    
            return null;
        }

        public ExcelLink GetExcelLink(string Name)
        {
            foreach (ExcelLink myVar in this.List)
            {
                if (myVar.ClassName.ToLower() == Name.ToLower()) return myVar;
            }
            return null;
        }

        public bool Contains(string Name)
        {
            foreach (ExcelLink myVar in this.List)
            {
                if (myVar.ClassName.ToLower() == Name) return true;
            }
            return false;
        }


        public ExcelLink this[int index]
        {
            get { return (ExcelLink)this.List[index]; }
        }

        public int IndexOf(int WorkSheetIndex)
        {
            foreach (ExcelLink myVar in this.List)
            {
                int CourseIndex = myVar.GetCourseIndex(WorkSheetIndex);
                if (CourseIndex!=-1) return CourseIndex;
            }
            return -1;
        }

        public void Remove(String Name)
        {
            foreach (ExcelLink myVar in this.List)
            {
                if (myVar.ClassName == Name)
                {
                    this.List.Remove(myVar);
                    break;
                }
            }
        }

        public new void RemoveAt(int index)
        {
            this.List.RemoveAt(index);
        }

    }
}
