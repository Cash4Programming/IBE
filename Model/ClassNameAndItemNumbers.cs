using System.Collections;


namespace InstructorBriefcaseExtractor.Model
{
    public class ClassNameAndItemNumbers : CollectionBase
    {
        // Creates an empty collection.
        public ClassNameAndItemNumbers()
        {
        }

        public ClassNameAndItemNumber Add(string ClassItemNumber)
        {
            ClassNameAndItemNumber value = new ClassNameAndItemNumber(ClassItemNumber);
            this.List.Add(value);
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

        public ClassNameAndItemNumber this[string ClassItemNumber]
        {
            get
            {
                foreach (ClassNameAndItemNumber myVar in this.List)
                {
                    if (myVar.ClassClassNameAndItemNumber == ClassItemNumber) return myVar;
                }
                return null;
            }
        }

        public ClassNameAndItemNumber this[int index]
        {
            get { return (ClassNameAndItemNumber)this.List[index]; }
        }

        public int IndexOf(string ClassItemNumber)
        {
            foreach (ClassNameAndItemNumber myVar in this.List)
            {
                if (myVar.ClassClassNameAndItemNumber == ClassItemNumber) return this.List.IndexOf(myVar);
            }
            return -1;
        }

        public void Remove(string ClassItemNumber)
        {
            foreach (ClassNameAndItemNumber myVar in this.List)
            {
                if (myVar.ClassClassNameAndItemNumber == ClassItemNumber)
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

    }    // class ClassNameAndItemNumber    
}
