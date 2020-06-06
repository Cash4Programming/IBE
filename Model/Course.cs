using System;
using InstructorBriefcaseExtractor.Utility;


namespace InstructorBriefcaseExtractor.Model
{
    // This class contains the information about a course and a list of all it's students
    public class Course
    {
        private readonly bool SaveOptionHeader = true;
        private string JavaStudentListVariable
        {
            get { return "new stuLine("; } 
        }

        private string JavaWaitListVariable
        {
            get { return "new classLn("; }
        }

        //private readonly HeaderColumn Phone;

        public Course(string RawHTML, string InstructorName, string CourseName, 
                      string ItemNumber, string YRQ, string QuarterName, ExcelClassSettings ExcelClassSettings)
        {
            SchoolHeaders = new OptionHeaders();
            StrInstructorName = InstructorName;
            StrCourseName = CourseName;
            StrItemNumber = ItemNumber;
            StrYRQ = YRQ;            
            StrQuarterName = QuarterName;
            // Set to export
            blnExport = true;

            if (RawHTML != "")
            {
                ExtractValue EV = new ExtractValue();
                StrTitle = EV.FromJavaScript(RawHTML, "title", "'", "'");
                StrSection = EV.FromJavaScript(RawHTML, "sect", "'", "'");
                StrCredits = EV.FromJavaScript(RawHTML, "credit", "'", "'");

                StrRoom = EV.FromJavaScript(RawHTML, "room", "'", "'");
                StrStartTime = EV.FromJavaScript(RawHTML, "strtTime", "'", "'");
                StrEndTime = EV.FromJavaScript(RawHTML, "endTime", "'", "'");
                StrDays = EV.FromJavaScript(RawHTML, "days", "'", "'");
                StrStartDate = EV.FromJavaScript(RawHTML, "strtDate", "'", "'");
                StrTenthDate = EV.FromJavaScript(RawHTML, "tenDay", "'", "'");

                StrOpt1Header = EV.FromJavaScript(RawHTML, "optHead1", "'", "'");
                StrOpt2Header = EV.FromJavaScript(RawHTML, "optHead2", "'", "'");
                StrOpt3Header = EV.FromJavaScript(RawHTML, "optHead3", "'", "'");
                StrOptWlHeader = EV.FromJavaScript(RawHTML, "optWLHead", "'", "'");

                if (SaveOptionHeader)
                {
                    ExcelClassSettings.HeaderNames.Header1 = StrOpt1Header;
                    ExcelClassSettings.HeaderNames.Header2 = StrOpt2Header;
                    ExcelClassSettings.HeaderNames.Header3 = StrOpt3Header;
                    ExcelClassSettings.HeaderNames.WaitListHeader = StrOptWlHeader;
                    SaveOptionHeader = false;
                }

                int Count = Convert.ToInt32(EV.FromJavaScript(RawHTML, "totStu", "'", "'"));
                myStudents = new Student[Count];

                Count = Convert.ToInt32(EV.FromJavaScript(RawHTML, "totStuWl", "'", "'"));
                myWaitlist= new Student[Count];

                // Extract all students from RawHTML
                ProcessStudentsFromHTML(RawHTML, ExcelClassSettings.HeaderNames);
            }
        }

        private void ProcessStudentsFromHTML(string RawHTML, OptionHeaders HeaderNames)
        {
            RawHTML = RawHTML.Replace("<BR>", " ").Replace("<BR/>", " ").Replace("<br>", " ").Replace("<br/>", " ");
            int HTMLStart = RawHTML.IndexOf(JavaStudentListVariable);
            int Count = 0;
            string temp;

            // Class
            // function stuLine( sid, name, enrCredit, gr, opt1, opt2, opt3 )
            do
            {
                int StudentStart = RawHTML.IndexOf("(", HTMLStart);
                int StudentEnd = RawHTML.IndexOf(");", StudentStart+45); // added ; on 12-16-2019 to prevent the ending to be accepted as part of the name
                string Student = RawHTML.Substring(StudentStart, StudentEnd - StudentStart);
                //stuLine( sid, name, enrCredit, gr, Opt1, Opt2, Opt3 )
                StudentStart = Student.IndexOf("'") + 1;
                StudentEnd = Student.IndexOf("'", StudentStart + 1);
                myStudents[Count] = new Student()
                {
                    SID = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim(),
                    Opt1Header = HeaderNames.Header1,
                    Opt2Header = HeaderNames.Header2,
                    Opt3Header = HeaderNames.Header3,

                };

                StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                StudentEnd = Student.IndexOf("'", StudentStart + 1);
                myStudents[Count].ExtractedName = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();

                StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                StudentEnd = Student.IndexOf("'", StudentStart + 1);
                myStudents[Count].EnrolledCredit = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();

                StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                StudentEnd = Student.IndexOf("'", StudentStart + 1);
                temp = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();
                if (temp != "&nbsp")
                {
                    myStudents[Count].gr = temp;
                }
                else
                {
                    myStudents[Count].gr = "";
                }

                StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                StudentEnd = Student.IndexOf("'", StudentStart + 1);
                temp = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();
                if (temp != "&nbsp")
                {
                    myStudents[Count].Opt1 = temp;
                }
                else
                {
                    myStudents[Count].Opt1 = "";
                }

                StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                StudentEnd = Student.IndexOf("'", StudentStart + 1);
                temp = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();
                if (temp != "&nbsp")
                {
                    myStudents[Count].Opt2 = temp;
                }
                else
                {
                    myStudents[Count].Opt2 = "";
                }

                StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                StudentEnd = Student.IndexOf("'", StudentStart + 1);
                // fixed change in HTML as the studetn email has been added.
                if ((StudentEnd - StudentStart) > 0)
                {
                    temp = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();

                    if (temp != "&nbsp")
                    {
                        myStudents[Count].Opt3 = temp;
                        myStudents[Count].Email = temp;
                    }
                    else
                    {
                        myStudents[Count].Opt3 = "";
                    }
                }
                Count += 1;
                HTMLStart = RawHTML.IndexOf(JavaStudentListVariable, HTMLStart + 1);
            } while (HTMLStart > 0);

            // Waitlist
            //  function classLn( sid, name, status,  datetime, opt3 )
            HTMLStart = RawHTML.IndexOf(JavaWaitListVariable);
            Count = 0;
            
            if (HTMLStart > 0)
            {
                do
                {
                    int StudentStart = RawHTML.IndexOf("(", HTMLStart);
                    int StudentEnd = RawHTML.IndexOf(")", StudentStart);
                    string Student = RawHTML.Substring(StudentStart, StudentEnd - StudentStart);
                    //stuLine( sid, name, W-Listed, Date, email)
                    StudentStart = Student.IndexOf("'") + 1;
                    StudentEnd = Student.IndexOf("'", StudentStart + 1);
                    myWaitlist[Count] = new Student()
                    {
                        SID = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim(),
                        Opt3WaitlistHeader = HeaderNames.WaitListHeader
                    };

                    StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                    StudentEnd = Student.IndexOf("'", StudentStart + 1);
                    myWaitlist[Count].ExtractedName = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();

                    StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                    StudentEnd = Student.IndexOf("'", StudentStart + 1);
                    myWaitlist[Count].WaitlistStatus = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();

                    StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                    StudentEnd = Student.IndexOf("'", StudentStart + 1);
                    myWaitlist[Count].WaitlistDate = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();

                    StudentStart = Student.IndexOf("'", StudentEnd + 1) + 1;
                    StudentEnd = Student.IndexOf("'", StudentStart + 1);
                    // fixed change in HTML as the student email has been added.
                    if ((StudentEnd - StudentStart) > 0)
                    {
                        temp = Student.Substring(StudentStart, StudentEnd - StudentStart).Trim();

                        if (temp != "&nbsp")
                        {
                            myWaitlist[Count].Opt3Waitlist = temp;
                            myWaitlist[Count].Email = temp;
                        }
                        else
                        {
                            myWaitlist[Count].Opt3 = "";
                        }
                    }
                    Count += 1;
                    HTMLStart = RawHTML.IndexOf(JavaWaitListVariable, HTMLStart + 1);
                } while (HTMLStart > 0);
            }
        }

        public OptionHeaders SchoolHeaders;

        private Student[] myWaitlist;
        public Student[] Waitlist
        {
            get { return myWaitlist; }
            set { myWaitlist = value; }
        }

        private Student[] myStudents;
        public Student[] Students
        { 
            get { return myStudents; }
            set { myStudents = value; }
        }

        private string StrCredits;
        public string Credits
        {
            get { return StrCredits; }
            set { StrCredits = value; }
        }

        private string StrDays;
        public string Days
        {
            get { return StrDays; }
            set { StrDays = value; }
        }

        public string DiskName(bool UseUnderScore)
        {
            string retval = StrQuarterName + " " + StrCourseName + "-" + StrSection + " (" + StrItemNumber + ")";
            if (UseUnderScore) { retval = "_" + retval; }
            return retval;
        }

        public string Display
        {
            get { return StrItemNumber + " " + StrCourseName; }
        }

        private string StrEndTime;
        public string EndTime
        {
            get { return StrEndTime; }
            set { StrEndTime = value; }
        }

        private bool blnExport;
        public bool Export
        {
            get { return blnExport; }
            set { blnExport = value; }
        }

        private string StrRoom;
        public string Room
        {
            get { return StrRoom; }
            set { StrRoom = value; }
        }

        private string StrSection;
        public string Section
        {
            get { return StrSection; }
            set { StrSection = value; }
        }

        private string StrStartDate;
        public string StartDate
        {
            get { return StrStartDate; }
            set { StrStartDate = value; }
        }

        private string StrTenthDate;
        public string TenthDate
        {
            get { return StrTenthDate; }
            set { StrTenthDate = value; }
        }

        private string StrStartTime;
        public string StartTime
        {
            get { return StrStartTime; }
            set { StrStartTime = value; }
        }

        public string SheetName()
        {
            //  modified 2018-02-06 added StartTime and the space
            return StrStartTime + "          " + StrCourseName + " (" + StrItemNumber + ")";
        }

        private string StrTitle;
        public string Title
        {
            get { return StrTitle; }
            set { StrTitle = value; }
        }

        private string StrInstructorName;
        public string InstructorName
        {
            get { return StrInstructorName; }
            set { StrInstructorName = value; }
        }

        private string StrItemNumber;
        public string ItemNumber
        {
            get { return StrItemNumber; }
            set { StrItemNumber = value; }
        }

        private string StrCourseName;
        public string Name
        {
            get { return StrCourseName; }
            set { StrCourseName = value; }
        }
        
        private string StrQuarterName;
        public string QuarterName
        {
            get { return StrQuarterName; }
            set { StrQuarterName = value; }
        }

        private string StrYRQ;
        public string YRQ
        {
            get { return StrYRQ; }
            set { StrYRQ = value; }
        }

        private string StrOpt1Header;
        public string Opt1Header
        {
            get { return StrOpt1Header; }
            set { StrOpt1Header = value; }
        }

        private string StrOpt2Header;
        public string Opt2Header
        {
            get { return StrOpt2Header; }
            set { StrOpt2Header = value; }
        }

        private string StrOpt3Header;
        public string Opt3Header
        {
            get { return StrOpt3Header; }
            set { StrOpt3Header = value; }
        }

        
        private string StrOptWlHeader;
        public string OptWaitlistHeader
        {
            get { return StrOptWlHeader; }
            set { StrOptWlHeader = value; }
        }


    }

}
