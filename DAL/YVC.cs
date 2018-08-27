using System;
using System.Collections.Generic;
using System.Text;

using InstructorBriefcaseExtractor.Model;

namespace InstructorBriefcaseExtractor.DAL
{
    public class YVC : College
    {
        private readonly College CollegeInfo = null; 

        public static string Key = "YVC";

        public YVC()
        {
            CollegeInfo = new College
            {
                // Set the information for the college            
                FullName = "Yakima Valley Community College",
                Abbreviation = YVC.Key,
                HasBeenVerified = true,
                EmployeeIDprompt = "865 Number",
                WebAssignInstitution = "yvcc.cc.wa",

                // Override any neded changes here
                // CISLogin
                RequestURL = "https://www.ctc.edu/~yvccwts/wts/ibc/"
            };
        }

        public College GetCollege
        {
            get { return CollegeInfo; }
        }
    }
    
}
