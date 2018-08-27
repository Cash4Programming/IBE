using System;
using System.Collections;
//using System.Collections.Generic;
using System.Text;

namespace InstructorBriefcaseExtractor.DAL
{
    public class Colleges : CollectionBase
    {
        public Colleges()
        {
            this.List.Add((new YVC()).GetCollege);
            this.List.Add((new GRC()).GetCollege);
        }

        public College FindCollege(string Abbreviation)
        {
            foreach (College Mvar in this.List)
            {
                if (Mvar.Abbreviation == Abbreviation)
                {
                    return Mvar;
                }
            }
            return null;
        }
    }
}
