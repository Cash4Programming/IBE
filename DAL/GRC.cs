using System;
using System.Collections.Generic;
using System.Text;

using InstructorBriefcaseExtractor.Model;

namespace InstructorBriefcaseExtractor.DAL
{
    public class GRC : College
    {
        private readonly College CollegeInfo = null;

        public static string Key = "GRC";

        public GRC()
        {
            CollegeInfo = new College
            {
                // Set the information for the college            
                FullName = "Green River College",
                Abbreviation = GRC.Key,
                HasBeenVerified = false,
                EmployeeIDprompt = "Employee ID Number",
                //https://docs.google.com/viewer?url=http://www.greenriver.edu/studentemail/instructions/student-email-first-time-login-instructions.pdf#zoom=100,0%2520,0&embedded=true&chrome=true
                WebAssignInstitution = "grcc.cc.wa",

                // Override any neded changes here
                // CISLogin
                RequestURL = "http://www.greenriver.edu/ibc/waci100.html"
            };
        }

        public College GetCollege
        {
            get { return CollegeInfo; }
        }
    }

}
