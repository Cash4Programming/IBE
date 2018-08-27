using System;
using System.Text;

using InstructorBriefcaseExtractor.Model;


namespace InstructorBriefcaseExtractor.DAL
{
    public class College 
    {
        public static string CollegeAbbreviationKey = "CollegeAbbreviation";

        public College()
        { }      

        // For selecting the college the user wants
        public bool HasBeenVerified { get; set; }  // Settings that are known to work
        public string FullName      { get; set; }  // Full Name of the college
        public string Abbreviation  { get; set; }  // Abbreviation  of the college
        public string EmployeeIDprompt { get; set; }  // label for form
        public string DropDownBoxName => this.FullName + " (" + this.Abbreviation + ")";

        // Export Specific options
        public string WebAssignInstitution { get; set; }  // college code for webassign
        public string RequestURL { get; set; }
    }
}  